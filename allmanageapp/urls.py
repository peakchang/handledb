from django.urls import path

from allmanageapp import views

appname = 'allmanage'
urlpatterns = [
    path('', views.super_manage, name='supermanage'),
]