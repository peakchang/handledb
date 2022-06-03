from django.contrib.auth import authenticate, login
from django.http import HttpResponse, HttpResponseRedirect
from django.shortcuts import render, redirect

# Create your views here.
from django.urls import reverse_lazy

from accountapp.forms import LoginForm
from accountapp.models import User


def testpage(request):
    return render(request, 'accountapp/testpage.html', {'test': 'test'})


def signup(request):
    if request.method == 'POST':
        if request.POST['password1'] != request.POST['password2']:
            error_message = {'error_nopass': '비밀번호가 일치하지 않습니다. 비밀번호를 확인해주세요'}
            return render(request, 'accountapp/signup.html', error_message)
        else:
            try:
                User.objects.create_user(username=request.POST['username'], password=request.POST['password1'], nickname=request.POST['nickname'], )
            except:
                error_message = {'error_noid': '이미 가입된 아이디 또는 닉네임 입니다.'}
                return render(request, 'accountapp/signup.html', error_message)

        user = authenticate(username=request.POST['username'], password=request.POST['password1'])
        login(request, user)
        return HttpResponseRedirect(reverse_lazy('home'))

    return render(request, 'accountapp/signup.html')


def signin(request):
    if request.method == 'POST':
        form = LoginForm(request.POST)
        if form.is_valid():  # form 검증
            username = form.cleaned_data['username']  # form에서 data 가져오기
            password = form.cleaned_data['password']
            user = authenticate(username=username, password=password)
            if user is not None:
                # 참일경우
                login(request, user)
                return HttpResponseRedirect(reverse_lazy('home'))

        # 유효성 검증에 실패했을경우
        return render(request, 'accountapp/login.html', {'error': '아이디와 비밀번호를 확인해주세요'})
    else:
        form = LoginForm()
        return render(request, 'accountapp/login.html')