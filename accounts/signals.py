from django.db.models.signals import post_save
from django.dispatch import receiver
from django.contrib.auth.models import User

from .models import UserProfile


@receiver(post_save, sender=User)
def create_user_profile(sender, instance, created, **kwargs):
    """
    Automatically create a UserProfile whenever
    a new Django User is created.
    """

    if created:
        UserProfile.objects.create(
            user=instance,
            employee_name=instance.first_name or instance.username,
            employee_id=f"EMP{instance.id:04}",
            department="IT Automation",
            role="RFQ",
        )


@receiver(post_save, sender=User)
def save_user_profile(sender, instance, **kwargs):
    """
    Save profile whenever User is saved.
    """

    if hasattr(instance, "userprofile"):
        instance.userprofile.save()