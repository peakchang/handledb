from django.db import models


# Create your models here.
class AllManage(models.Model):
    sample_excel_file = models.FileField(upload_to="sample/", null=True, blank=True)
    sample_image = models.ImageField(upload_to='sample/', null=True, blank=True)
    end_time = models.DateTimeField(auto_now_add=True, null=True)
    now_buistatus = models.CharField(max_length=20, null=True, default='')