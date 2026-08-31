from django.urls import path
from .views import (
    login_view,
    logout_view,
    user_management,
    create_user,
    edit_user,
    toggle_user_status,
    reset_password,
    delete_user,
)

urlpatterns = [
    # Authentication
    path("", login_view, name="login"),
    path("logout/", logout_view, name="logout"),

    # User Management (Admin Only)
    path("users/", user_management, name="user_management"),
    path("users/create/", create_user, name="create_user"),
    path("users/edit/<int:user_id>/", edit_user, name="edit_user"),
    path("users/status/<int:user_id>/", toggle_user_status, name="toggle_user_status"),
    path("users/reset-password/<int:user_id>/",reset_password,name="reset_password",),
    path("users/delete/<int:user_id>/",delete_user, name="delete_user"),

]