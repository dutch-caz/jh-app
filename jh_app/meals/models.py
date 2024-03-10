from django.db import models


MEAL_TYPE = (
    ('breakfast', 'Breakfast'), 
    ('lunch', 'Lunch'),
    ('dinner', 'Dinner'),
)

PERSON_TYPE = (
    ('crew', 'Crew'),
    ('staff', 'Staff'),
    ('intercessor', 'Intercessor'),
    ('guest', 'Guest'),
)

class enableSignUp(models.Model):
    signUpEnabled = models.BooleanField('Enable Sign Ups?', default=False)
    
    def __str__(self):
        return "Enable Sign Ups Here"

class attendee(models.Model):
    name = models.CharField('Name', max_length=30)
    person_type = models.CharField('Person Type', choices=PERSON_TYPE, default='crew', max_length=30)
    half_plates = models.FloatField('Half Plates', default=0.0)
    gluten_free = models.BooleanField('Gluten Free', default=False)
    dairy_free = models.BooleanField('Dairy Free', default=False)

    def __str__(self):
        return self.name
    
    
class meal(models.Model):
    name = models.CharField('Meal Name', max_length=40)
    description = models.TextField('Meal Description', blank=True)
    meal_type = models.CharField('Meal Type', choices=MEAL_TYPE, default='breakfast', max_length=30)
    start_time = models.DateTimeField('Start Time and Date')
    end_time = models.DateTimeField('End Time and Date')
    latitude = models.DecimalField('Set Latitude', max_digits=40, decimal_places=3)
    longitude = models.DecimalField('Set Longitude', max_digits=40, decimal_places=3)
    people_attending = models.ManyToManyField('attendee', through='attendance', blank=True, null=True)
    
    def __str__(self):
        return self.name
    
    
class attendance(models.Model):
    attendee = models.ForeignKey(attendee, on_delete=models.CASCADE, blank=True, null=True)
    meal = models.ForeignKey(meal, on_delete=models.CASCADE, blank=True, null=True, default=None)
    present = models.BooleanField('Present', default=False)
    
    class Meta():
        unique_together = ('attendee', 'meal')
    
    def __str__(self):
        return f"{self.attendee.name} - {self.meal.name}"
    
        