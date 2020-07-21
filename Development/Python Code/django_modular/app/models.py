from django.db import models

# Create your models here.
class senders(models.Model):
    name = models.CharField(max_length = 300)
    address = models.CharField(max_length = 300)
    sender_total_mails = models.IntegerField(default = 0)
    sender_reply_count = models.IntegerField()
    sender_delete_count = models.IntegerField(default = 0)
    sender_opened_count = models.IntegerField(default = 0)
    sender_total_count = models.IntegerField(default = 0)
    sender_importance = models.FloatField()

    def __str__(self):
        return self.name+" "+self.address

class email(models.Model):
    id_mail = models.CharField(max_length = 300)
    sender_address = models.CharField(max_length = 300)
    subject = models.CharField(max_length = 300)
    body = models.CharField(max_length = 1000000)
    app_open_time = models.CharField(max_length = 300)
    mail_open_time = models.CharField(max_length = 300)
    read = models.BooleanField(default = False)
    isreplied_count = models.IntegerField(default = 0)
    first_reply_time = models.CharField(max_length = 100)
    body_spams_score = models.FloatField(default = 0.0)
    type_of = models.CharField(max_length = 20, default = None)
    received_time = models.CharField(max_length = 100, default = None)
    completeflag = models.BooleanField(default = False)
    sender_name = models.CharField(max_length = 300, default = "sender")
    