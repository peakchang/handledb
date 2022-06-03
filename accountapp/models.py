from django.contrib.auth.models import AbstractUser
from django.db import models

# Create your models here.

RATE_CHOICE = (('A', '마스터'), ('B', '관리자'), ('C', '부관리자'), ('D', '직원'))
STATUS_CHOICE = (('Y', '활동'), ('N', '비활동'))


class User(AbstractUser):
    rete = models.CharField(choices=RATE_CHOICE, max_length=20, null=False, default='D')
    nickname = models.CharField(max_length=20, null=False, default='', unique=True)
    status = models.CharField(choices=STATUS_CHOICE, max_length=5, null=False, default='Y')
