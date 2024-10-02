# metrics_app/models.py
from django.db import models

class Metric(models.Model):
    name = models.CharField(max_length=100)
    description = models.TextField()
    criteria = models.JSONField()  # Use JSONField to store criteria as a list

    def __str__(self):
        return self.name

class UserSubscription(models.Model):
    user = models.OneToOneField('auth.User', on_delete=models.CASCADE)
    start_date = models.DateTimeField(auto_now_add=True)
    end_date = models.DateTimeField()
    active = models.BooleanField(default=True)
