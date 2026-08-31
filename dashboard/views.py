from django.shortcuts import render
from django.views.decorators.cache import never_cache
from accounts.decorators import role_required

from django.contrib.auth.models import User
from accounts.models import UserProfile


@never_cache
@role_required(["ADMIN", "RFQ", "GRDB", "STAT"])
def dashboard(request):

    profile = request.user.userprofile

    context = {
        # Logged-in user info
        "role": profile.role,
        "employee_name": profile.employee_name,

        # Dashboard Statistics
        "total_users": UserProfile.objects.count(),
        "active_users": User.objects.filter(is_active=True).count(),
        "disabled_users": User.objects.filter(is_active=False).count(),

        "admin_users": UserProfile.objects.filter(role="ADMIN").count(),
        "rfq_users": UserProfile.objects.filter(role="RFQ").count(),
        "grdb_users": UserProfile.objects.filter(role="GRDB").count(),
        "stat_users": UserProfile.objects.filter(role="STAT").count(),
    }

    return render(request, "dashboard.html", context)