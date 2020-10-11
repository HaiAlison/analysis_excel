import pandas as pd
from django.shortcuts import render,redirect
from django.http import HttpResponse, HttpResponseRedirect
from .forms import ImportFile
from django.core.files.storage import FileSystemStorage
from django.contrib import messages
from .ABCAnalysis import *
from .XYZAnalysis import *
from .ToExcel import *
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from rest_framework import viewsets
from .models import Import
from .serializers import ImportSerializer
# from rest_framework.response import Response
# from rest_framework.views import APIView
# from rest_framework.parsers import MultiPartParser, FormParser


from io import BytesIO
import xlsxwriter
import datetime


# class ImportFile(viewsets.ModelViewSet):
#     queryset = Import.objects.all()
#     serializer_class = ImportSerializer
# class InsertFile(APIView):
#     parser_classes = (MultiPartParser, FormParser)

#     def post(self, request, *args, **kwargs):
#         file_serializer = FileSerializer(data=request.data)
#         if file_serializer.is_valid():
#             uploaded_file = request.FILES['link_to_specs']
#             fs = FileSystemStorage()
#             dat = fs.save(uploaded_file.name, uploaded_file)
#             return handle_upload_file(request,dat)
#         else:
#             return Response(f'Please insert file!', status=status.HTTP_400_BAD_REQUEST)



def home(request):
    if request.method == "POST":
        form = ImportFile(request.POST, request.FILES)
        if form.is_valid():
            form.save()
            uploaded_file = request.FILES['link_to_specs']
            fs = FileSystemStorage()
            dat = fs.save(uploaded_file.name, uploaded_file)
            return handle_upload_file(request,dat)
        else:
            messages.error(request,f'Please insert file!')
            form = ImportFile()
    return render(request, 'analysis/analysis.html', {'title': 'Home'})


def handle_upload_file(request,f):
    with BytesIO() as b:
        with pd.ExcelFile("media/"+f) as reader: 
            now = datetime.datetime.now().strftime("%m%d%y_%H%M%S")
            filename = 'result_'+str(now)
            header = ['STT', 'MÃ HÀNG', 'MÃ MẶT HÀNG', 'TÊN THUỐC', 'HOẠT CHẤT', 'HÀM LƯỢNG', 'ĐVT', 'TCKT', 'VEN', 'HÃNG SX', 'NƯỚC SX', 'NHÀ CUNG ỨNG', 'ĐƠN GIÁ', 'SỐ LƯỢNG', 'THÀNH TIỀN', 'NHÓM THUỐC', 'HỆ ĐIỀU TRỊ', 'NĂM']
            index = pd.read_excel(reader,sheet_name = 0).columns
            head = []
            for data in index:
                head.append(data)
            if(header == head):
                dataABC = pd.read_excel(reader, sheet_name=0)  
                dataXYZ = pd.read_excel(reader, sheet_name=1) 
                dataResultsABC = ABCVEN(dataABC)
                dataResultsXYZ = AnalysisXYZ(dataXYZ)
                wb = Workbook()
                ws = wb.active
                ws.title = "ABC_VEN_Data"
                for r in dataframe_to_rows(pd.concat(dataResultsABC), index=False, header=True):
                    ws.append(r)
                ABCAnalysisDisplay(wb, pd.concat(dataResultsABC))
                ws = wb.create_sheet(title="XYZ_Data")
                for r in dataframe_to_rows(dataResultsXYZ, index=False, header=True):
                    ws.append(r)
                XYZAnalysisDisplay(wb, dataResultsXYZ)
                b.seek(0)
                wb.save(b)
                messages.success(request,'successfully uploaded!')
                response = HttpResponse(b.getvalue(), content_type='application/vnd.ms-excel')
                response['Content-Disposition'] = 'attachment; filename={}.xlsx'.format(filename)
                return response
            else:
                print("Fail")
                messages.error(request,'Please insert correct format!')
                return render(request, 'blog/home.html', {'title': 'Home'})

        