# Generated by Django 5.0.2 on 2024-03-09 20:56

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('meals', '0007_remove_attendee_vegan'),
    ]

    operations = [
        migrations.AlterField(
            model_name='attendee',
            name='half_plates',
            field=models.FloatField(default=0.0, verbose_name='Half Plates'),
        ),
    ]