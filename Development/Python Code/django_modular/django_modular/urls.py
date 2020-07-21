"""django_modular URL Configuration

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/2.2/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""
from django.contrib import admin
from django.urls import path
from . import views
import app.views as second
urlpatterns = [
    path('admin/', admin.site.urls),
    path('index.html', views.main),
    path('sentiment.html', views.sentiment),
    path('addintoexcel.html', views.toexcel),
    path('priority.html', views.priority),
    path('excel.html', views.excel),
    path('developers.html', views.develop),
    path('new', views.updatesentiment),
    path('customize.html', views.custom),
    path('second.html', second.replied_till_addin),
    path('secondnext.html', second.reply_update),
    path('third.html', second.rating_updater),
    path('imp.html', second.sender_importance),
    path('rank.html', second.sender_rank),
]
