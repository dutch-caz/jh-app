from django.contrib import admin
from django.urls import path, include
from django.conf import settings
from django.conf.urls.static import static

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', include('meals.urls')),
    
]
urlpatterns += static(settings.MEDIA_URL, document_root = settings.MEDIA_ROOT)
urlpatterns += static(settings.STATIC_URL, document_root=settings.STATIC_ROOT)


# Edit the admin area!

admin.site.site_header = "JH Ranch Administration"
admin.site.site_title = "JH Ranch"
admin.site.enable_nav_sidebar = True