import os

from django.shortcuts import render


# Create your views here.
from allmanageapp.models import AllManage


def super_manage(request):
    if request.method == 'POST':
        temp_super = AllManage.objects.last()
        sample_exl = request.FILES.get('sample_exl')
        sample_img = request.FILES.get('sample_img')
        end_time = request.POST.get('end_date')
        nowstatus = request.POST.get('nowstatus')
        if sample_exl:
            temp_super.sample_excel_file = sample_exl
            if temp_super.sample_excel_file:
                os.remove(temp_super.sample_excel_file.path)
        if sample_img:
            temp_super.sample_image = sample_img
            if temp_super.sample_img:
                os.remove(temp_super.sample_img.path)
        if end_time:
            temp_super.end_time = end_time
        temp_super.now_buistatus = nowstatus
        temp_super.save()

    set_supaermanage = AllManage.objects.last()
    print(set_supaermanage)
    return render(request, 'allmanageapp/allmanage.html')