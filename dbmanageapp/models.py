from django.core.validators import RegexValidator
from django.db import models

ADV_CHOICE = (('cpa', 'CPA'), ('cpc', 'CPC'), ('cps', 'CPS'), ('cpv', 'CPV'), ('cpm', 'CPM'))
APPROV_CHOICE = (('Y', '거래'), ('N', '미거래'))
THEME_CHOICE = (('light', '라이트'), ('dark', '다크'))
# Create your models here.

class DbSetting(models.Model):
    ds_status = models.CharField(max_length=255, null=True, blank=True)
    ds_statusbase = models.CharField(max_length=20, null=True, blank=True, default='')
    ds_overlapallow = models.CharField(max_length=255, null=True, blank=True, default='')
    logo_image = models.ImageField(upload_to='setimage/', null=True, blank=True)
    company_name = models.CharField(max_length=255, null=True)
    theme_status = models.CharField(choices=THEME_CHOICE, max_length=5, null=False, default='dark')

class MarketingList(models.Model):
    mk_company = models.CharField(max_length=30, null=False)
    mk_name = models.CharField(max_length=10, null=True)
    phoneNumberRegex = RegexValidator(regex=r'^01([0|1|6|7|8|9]?)-?([0-9]{3,4})-?([0-9]{4})$')
    mk_phone = models.CharField(validators=[phoneNumberRegex], max_length=11, null=True)
    mk_advtype = models.CharField(choices=ADV_CHOICE, max_length=5, null=True, default='cpa')
    mk_url = models.CharField(max_length=255, null=True)
    mk_memo = models.TextField(null=True)
    mk_status = models.CharField(choices=APPROV_CHOICE, max_length=5, null=False, default='Y')
    mk_date = models.DateTimeField(auto_now_add=True, null=True)

    # def __str__(self):
    #     return self.mk_company

class UploadDbName(models.Model):
    dbn_mkname = models.ForeignKey(MarketingList, on_delete=models.SET_NULL, null=True, related_name='mkdname')
    dbn_name = models.CharField(max_length=255, null=False)
    dbn_price = models.IntegerField(null=False)
    dbn_memo = models.CharField(max_length=255, null=True)
    dbn_date = models.DateTimeField(auto_now_add=True, null=True)

    # def __str__(self):
    #     return self.dbn_name


class UploadDb(models.Model):
    db_name = models.ForeignKey(UploadDbName, on_delete=models.SET_NULL, null=True, related_name='dbname')
    db_mkname = models.ForeignKey(MarketingList, on_delete=models.SET_NULL, null=True, related_name='dbmkname')
    db_phone = models.CharField(max_length=20, null=False)
    db_member = models.CharField(max_length=20, null=False)
    db_age = models.CharField(max_length=20, null=True, default='')
    db_sex = models.CharField(max_length=20, null=True, default='')
    db_inv = models.CharField(max_length=20, null=True, default='')
    db_manager = models.CharField(max_length=20, null=True)
    db_manager_nick = models.CharField(max_length=20, null=True)
    db_status = models.CharField(max_length=20, null=True, default='')
    db_paidprice = models.IntegerField(null=True, default='0', blank=True)
    db_paidstatus = models.CharField(max_length=5, null=True, default='N')
    db_lastpaiddate = models.DateTimeField(auto_now_add=True, null=True)
    db_date = models.DateTimeField(auto_now_add=True, null=True)
    # db_divdate = models.DateTimeField(auto_now_add=True, null=True)



    # def __str__(self):
    #     return self.db_member


class DbMemo(models.Model):
    dm_chkdb = models.ForeignKey(UploadDb, on_delete=models.CASCADE, null=False, related_name='dmchkdb', default='')
    dm_manager = models.CharField(max_length=50, null=True, default='')
    dm_memos = models.CharField(max_length=2550, null=False, default='')
    dm_date = models.DateTimeField(auto_now_add=True, null=False)


class PaidList(models.Model):
    pl_chkdb = models.ForeignKey(UploadDb, on_delete=models.CASCADE, null=False, related_name='plchkdb', default='')
    pl_paidprice = models.IntegerField(null=False, default='0', blank=False)
    pl_paiddate = models.DateTimeField(auto_now_add=True, null=False)


class TestField(models.Model):
    td_name = models.CharField(max_length=15)
    td_mkname = models.ManyToManyField(MarketingList, related_name='testfield')


class ChkDbReset(models.Model):
    chk_dbreset = models.CharField(max_length=15)
    chk_dbtime = models.DateTimeField(auto_now_add=True, null=False)