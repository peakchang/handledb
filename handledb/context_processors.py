# context_processors.py
from django.utils import timezone

from dbmanageapp.models import DbSetting


def message_processor(request):

    try:
        base_item = DbSetting.objects.last()
        base_logo = base_item.logo_image
        base_company_name = base_item.company_name
        base_time = timezone.now().strftime ("%m%d%M%S")
        base_theme = base_item.theme_status
    except:
        return {}
    return {'base_logo' : base_logo, 'base_company_name': base_company_name,'base_time': base_time, 'base_theme': base_theme}