from django.contrib import admin
from .models import UserProfile


@admin.register(UserProfile)
class UserProfileAdmin(admin.ModelAdmin):

    # Columns visible in admin list
    list_display = (
        "employee_name",
        "employee_id",
        "department",
        "role",
        "user",
        "created_at",
    )

    # Search bar
    search_fields = (
        "employee_name",
        "employee_id",
        "user__username",
        "user__email",
    )

    # Filters on right side
    list_filter = (
        "role",
        "department",
    )

    # Sort by latest created users
    ordering = ("-created_at",)

    # Read-only fields
    readonly_fields = ("created_at",)