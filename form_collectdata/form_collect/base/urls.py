from django.urls import path
from . import views
urlpatterns = [
    path('', views.index, name='index'),
    path('hoiDongChuyenMon', views.hoiDongChuyenMon, name='hoiDongChuyenMon'),
    path('baoCaoTienDoL1', views.baoCaoTienDoL1, name='baoCaoTienDoL1'),
    path('baoCaoTienDoL2', views.baoCaoTienDoL2, name='baoCaoTienDoL2'),
    path('huongdan3', views.huongdan3, name='huongdan3'),
    path('canBoPhanBien', views.canBoPhanBien, name='canBoPhanBien'),
    path('get-students', views.get_students_view, name='get_students'),
    path('get-all-students', views.get_all_students_view, name='get_all_students'),
    path('get-all-councils', views.get_all_councils_view, name='get_all_councils'),
    # path('process_form_hd1/', views.process_form_hd1, name='process_form_hd1'),
    path('process_form_hd2/', views.process_form_hd2_new, name='process_form_hd2'),
    path('process_form_hd3/', views.process_form_hd3_new, name='process_form_hd3'),
    # path('process_form_hdcm/', views.process_form_hdcm, name='process_form_hdcm'),
    path('process_form_pb/', views.process_form_pb_new, name='process_form_pb'),
    path('process_form_hdcm_new/', views.process_form_hdcm_new, name='process_form_hdcm_new'),
    path('process_form_hd1_new/', views.process_form_hd1_new, name='process_form_hd1_new')
]