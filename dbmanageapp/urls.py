from django.urls import path

from dbmanageapp import views

app_name = 'dbmanage'
urlpatterns = [
    path('', views.dbmainpage, name='mainpage'),
    path('accountmanagement/', views.accountmanagement, name='accountmanagement'),
    path('markterlist/', views.markerlist, name='markerlist'),
    path('newdbup/', views.newdbup, name='newdbup'),
    path('divdb/', views.divdb, name='divdb'),
    path('alldblist/', views.alldblist, name='alldblist'),
    path('sale_st/', views.sale_st, name='sale_st'),
    path('detail_customer/<int:id>', views.detail_customer, name='detail_customer'),
    path('workajax/', views.workAjax, name='workajax'),

    path('emp_dblist/', views.emp_dblist, name='emp_dblist'),
    path('emp_dbstats/', views.emp_dbstats, name='emp_dbstats'),
    path('base_setting/', views.base_setting, name='base_setting'),
    path('marketer_stats/', views.marketer_stats, name='marketer_stats'),
    path('test_chk/', views.test_chk, name='test_chk'),
    path('status_stats/', views.status_stats, name='status_stats'),
    path('updatedb/', views.update_db, name='updatedb')



]