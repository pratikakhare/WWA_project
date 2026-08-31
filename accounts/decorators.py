from functools import wraps
from django.contrib.auth.decorators import login_required
from django.shortcuts import render
from django.contrib import messages


def role_required(allowed_roles):
    """
    Allow access only to users with one of the allowed roles.
    Example:
        @role_required(["ADMIN", "RFQ"])
    """

    def decorator(view_func):

        @login_required(login_url="login")
        @wraps(view_func)
        def wrapper(request, *args, **kwargs):

            # Check UserProfile exists
            profile = getattr(request.user, "userprofile", None)

            if profile is None:
                messages.error(request, "User profile is missing.")
                return render(request, "403.html", status=403)

            # Check role
            if profile.role not in allowed_roles:
                messages.error(
                    request,
                    "You do not have permission to access this page."
                )
                return render(request, "403.html", status=403)

            return view_func(request, *args, **kwargs)

        return wrapper

    return decorator