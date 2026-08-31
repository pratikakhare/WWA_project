
from django.shortcuts import render
from accounts.decorators import role_required
from django.views.decorators.cache import never_cache
# Create your views here.

@never_cache
@role_required(["ADMIN", "STATI"])
def stati_dashboard(request):
    return render(request, "stati/stati_dashboard.html") 