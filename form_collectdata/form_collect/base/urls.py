from django.urls import path
from . import views
urlpatterns = [
    path('',views.index,name='index'),
    path('form1',views.form1,name='form1'),
    path('form2',views.form2,name='form2')
]