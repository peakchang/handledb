from django.contrib import admin

# Register your models here.
from dbmanageapp.models import MarketingList, UploadDbName, UploadDb, DbSetting, DbMemo, TestField, PaidList

admin.site.register(MarketingList)

admin.site.register(UploadDbName)


# admin.site.register(UploadDb)
class Chk_updb(admin.ModelAdmin):
    list_display = ('db_phone', 'db_date')
admin.site.register(UploadDb,Chk_updb)

admin.site.register(DbSetting)

admin.site.register(DbMemo)

admin.site.register(TestField)

admin.site.register(PaidList)
