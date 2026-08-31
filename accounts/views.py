from django.shortcuts import render, redirect
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.models import User
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.views.decorators.http import require_POST
from django.http import HttpResponseForbidden
from django.shortcuts import get_object_or_404
from .models import UserProfile


# --------------------------------------------------
# Login
# --------------------------------------------------
def login_view(request):

    if request.user.is_authenticated:
        if hasattr(request.user, "userprofile"):
            return redirect("dashboard")

        logout(request)
        messages.warning(request, "Please login again.")

    if request.method == "POST":

        username = request.POST.get("username", "").strip()
        password = request.POST.get("password", "").strip()

        # Check if the user exists first
        try:
            existing_user = User.objects.get(username=username)

            # User exists but is disabled
            if not existing_user.is_active:
                messages.error(
                    request,
                    "Your account has been disabled. Contact the administrator."
                )
                return redirect("login")

        except User.DoesNotExist:
            pass

        # Authenticate active users
        user = authenticate(request, username=username, password=password)

        if user is not None:
            login(request, user)
            return redirect("dashboard")

        messages.error(request, "Invalid username or password.")

    return render(request, "login.html")


# --------------------------------------------------
# Logout
# --------------------------------------------------
@require_POST
@login_required(login_url="login")
def logout_view(request):
    logout(request)
    messages.success(request, "You have been logged out successfully.")
    return redirect("login")


# --------------------------------------------------
# User Management (Admin Only)
# --------------------------------------------------
@login_required(login_url="login")
def user_management(request):

    # Only ADMIN can access this page
    if request.user.userprofile.role != "ADMIN":
        return render(request, "403.html", status=403)

    users = UserProfile.objects.select_related("user").order_by("employee_name")

    # Dashboard counts
    active_users = users.filter(user__is_active=True).count()
    disabled_users = users.filter(user__is_active=False).count()

    context = {
        "users": users,
        "active_users": active_users,
        "disabled_users": disabled_users,
    }

    return render(
        request,
        "user_management.html",
        context,
    )


# --------------------------------------------------
# Create User (Admin Only)
# --------------------------------------------------
@login_required(login_url="login")
def create_user(request):

    # Only ADMIN can create users
    if request.user.userprofile.role != "ADMIN":
        return render(request, "403.html", status=403)

    if request.method == "POST":

        username = request.POST.get("username", "").strip()
        password = request.POST.get("password", "").strip()
        email = request.POST.get("email", "").strip().lower()
        name = request.POST.get("name", "").strip()
        emp_id = request.POST.get("employee_id", "").strip()
        department = request.POST.get("department")
        role = request.POST.get("role")

        # Validation
        if User.objects.filter(username=username).exists():
            messages.error(request, "Username already exists.")
            return redirect("create_user")

        if User.objects.filter(email=email).exists():
            messages.error(request, "Email already exists.")
            return redirect("create_user")

        if UserProfile.objects.filter(employee_id=emp_id).exists():
            messages.error(request, "Employee ID already exists.")
            return redirect("create_user")

        # Create Django User
        user = User.objects.create_user(
            username=username,
            password=password,
            email=email,
            first_name=name,
            is_active=True,
        )

        # Update the profile created automatically by signals
        profile = user.userprofile
        profile.employee_name = name
        profile.employee_id = emp_id
        profile.department = department
        profile.role = role
        profile.save()

        messages.success(request, f"{name} created successfully.")
        return redirect("user_management")

    return render(request, "create_user.html")


#EDIT only Admin
@login_required(login_url="login")
def edit_user(request, user_id):

    # Only ADMIN can edit users
    if request.user.userprofile.role != "ADMIN":
        return render(request, "403.html", status=403)

    profile = get_object_or_404(
        UserProfile.objects.select_related("user"),
        id=user_id
    )

    if request.method == "POST":

        username = request.POST.get("username").strip()
        email = request.POST.get("email").strip()
        name = request.POST.get("name").strip()
        department = request.POST.get("department")
        role = request.POST.get("role")

        # Check duplicate username
        if User.objects.filter(username=username).exclude(id=profile.user.id).exists():
            messages.error(request, "Username already exists.")
            return redirect("edit_user", user_id=user_id)

        # Check duplicate email
        if User.objects.filter(email=email).exclude(id=profile.user.id).exists():
            messages.error(request, "Email already exists.")
            return redirect("edit_user", user_id=user_id)

        # Update Django User
        profile.user.username = username
        profile.user.email = email
        profile.user.first_name = name
        profile.user.save()

        # Update Profile
        profile.employee_name = name
        profile.department = department
        profile.role = role
        profile.save()

        messages.success(request, "User updated successfully.")
        return redirect("user_management")

    return render(
        request,
        "edit_user.html",
        {"profile": profile}
    )
    
    

@login_required(login_url="login")
def toggle_user_status(request, user_id):

    # Only ADMIN can enable/disable users
    if request.user.userprofile.role != "ADMIN":
        return render(request, "403.html", status=403)

    profile = get_object_or_404(
        UserProfile.objects.select_related("user"),
        id=user_id
    )

    # Prevent admin from disabling themselves
    if profile.user == request.user:
        messages.error(request, "You cannot disable your own account.")
        return redirect("user_management")

    # Toggle active status
    profile.user.is_active = not profile.user.is_active
    profile.user.save()

    if profile.user.is_active:
        messages.success(
            request,
            f"{profile.employee_name} has been enabled."
        )
    else:
        messages.success(
            request,
            f"{profile.employee_name} has been disabled."
        )

    return redirect("user_management")




#Reset password

# --------------------------------------------------
# Reset Password (Admin Only)
# --------------------------------------------------
@login_required(login_url="login")
def reset_password(request, user_id):

    # Only ADMIN can reset passwords
    if request.user.userprofile.role != "ADMIN":
        return render(request, "403.html", status=403)

    profile = get_object_or_404(
        UserProfile.objects.select_related("user"),
        id=user_id
    )

    if request.method == "POST":

        new_password = request.POST.get("new_password", "").strip()
        confirm_password = request.POST.get("confirm_password", "").strip()

        # Validation
        if new_password != confirm_password:
            messages.error(request, "Passwords do not match.")
            return redirect("reset_password", user_id=user_id)

        if len(new_password) < 8:
            messages.error(
                request,
                "Password must contain at least 8 characters."
            )
            return redirect("reset_password", user_id=user_id)

        # Reset Password
        profile.user.set_password(new_password)
        profile.user.save()

        messages.success(
            request,
            f"Password reset successfully for {profile.employee_name}."
        )

        return redirect("user_management")

    return render(
        request,
        "reset_password.html",
        {"profile": profile}
    )



## delete user


from django.shortcuts import get_object_or_404

@login_required(login_url="login")
def delete_user(request, user_id):

    if request.user.userprofile.role != "ADMIN":
        return render(request, "403.html", status=403)

    profile = get_object_or_404(UserProfile, id=user_id)

    # Don't allow admin to delete themselves
    if profile.user == request.user:
        messages.error(request, "You cannot delete your own account.")
        return redirect("user_management")

    employee_name = profile.employee_name

    # Delete Django User (UserProfile deletes automatically)
    profile.user.delete()

    messages.success(request, f"{employee_name} deleted successfully.")
    return redirect("user_management")