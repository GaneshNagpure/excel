from django.urls import path
from . import views

urlpatterns = [
    path('upload/', views.upload_excel, name='upload_excel'),
    #path('bar_graph/', views.bar_graph, name='bar_graph'),
]
