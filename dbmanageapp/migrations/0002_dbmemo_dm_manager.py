# Generated by Django 4.0.4 on 2022-06-03 21:11

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('dbmanageapp', '0001_initial'),
    ]

    operations = [
        migrations.AddField(
            model_name='dbmemo',
            name='dm_manager',
            field=models.CharField(default='', max_length=50, null=True),
        ),
    ]
