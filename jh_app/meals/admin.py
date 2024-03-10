from django.contrib import admin
from .models import attendance
from .models import attendee
from .models import meal
from .models import enableSignUp


# Show models in admin area
admin.site.register(attendee)
admin.site.register(meal)
admin.site.register(attendance)
admin.site.register(enableSignUp)