from django.db import models
from django.contrib.auth.models import User


class UserProfile(models.Model):

    ROLE_CHOICES = [
        ("ADMIN", "Admin"),
        ("RFQ", "RFQ Team"),
        ("GRDB", "GRDB Team"),
        ("STATI", "STAT Team"),
    ]

    DEPARTMENT_CHOICES = [
        ("IT Automation", "IT Automation"),
        ("RFQ", "RFQ Team"),
        ("GRDB", "GRDB Team"),
        ("STATI", "STAT Team"),
    ]

    user = models.OneToOneField(
        User,
        on_delete=models.CASCADE,
        related_name="userprofile"
    )

    employee_name = models.CharField(max_length=100)
    employee_id = models.CharField(max_length=20, unique=True)

    department = models.CharField(max_length=50,default="IT Automation")
    role = models.CharField(max_length=20,choices=ROLE_CHOICES,default="RFQ")

    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.employee_name} ({self.role})"