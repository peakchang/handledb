# Generated by Django 4.0.4 on 2022-06-07 19:50

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('dbmanageapp', '0002_dbmemo_dm_manager'),
    ]

    operations = [
        migrations.AddField(
            model_name='dbsetting',
            name='ds_overlapallow',
            field=models.CharField(blank=True, default='', max_length=255, null=True),
        ),
    ]
