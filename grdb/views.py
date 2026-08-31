from django.shortcuts import render
from django.views.decorators.cache import never_cache
from accounts.decorators import role_required


@never_cache
@role_required(["ADMIN", "GRDB"])
def grdb_dashboard(request):
    return render(request, "grdb/grdb_dashboard.html")