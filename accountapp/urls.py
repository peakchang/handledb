from django.contrib.auth.views import LoginView, LogoutView
from django.urls import path

from accountapp import views

app_name = 'accountapp'

urlpatterns = [
    path('signup/', views.signup, name='signup'),
    path('login/', views.signin, name='login'),
    path('logout/', LogoutView.as_view(), name='logout'),
    path('logout/', views.testpage, name='testpage'),
]