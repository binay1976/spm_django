from django.urls import path
from . import views
from .views import upload_medha, home
from .views import extract_cms_ids

app_name = "spmApp"

urlpatterns = [
    path('', views.index, name='index'),
    path('home/', views.home, name='home'),
    path('upload-medha/', views.upload_medha, name='upload_medha'),
    path('upload-telpro-pdf/', views.upload_telpro_pdf, name='upload_telpro_pdf'),
    path('upload-laxvan/', views.upload_laxvan, name='upload_laxvan'),
    path('upload-quick-report/', views.upload_quick_report, name='upload_quick_report'),
    path('clean_temp_files/', views.clean_temp_files, name='clean_temp_files'),
    path('extract-cms-ids/', views.extract_cms_ids, name='extract_cms_ids'),
]

from django.conf import settings
from django.conf.urls.static import static

urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
