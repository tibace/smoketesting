import shutil

from django.core.files.storage import FileSystemStorage
from django.shortcuts import render
from django.http import HttpResponse, Http404
import datetime
import os
from django.conf import settings
from django.conf.urls.static import static
import subprocess
from subprocess import PIPE
import sys
import mimetypes
from .models import MultipleImage
from django.contrib import messages



def button(request):
    return render(request, 'geniusvoice.html')




def output(request):
    now = subprocess.run([sys.executable, "C:\\Users\mkanniah\\automation-v1.1\\automation\\st_mileposthelper.py", "print('')"],shell=False,stdout=PIPE)
    html = "<html><body> %s <br><br><a href='/'>Back<a></body></html>" % now
    if html.__contains__('returncode=1'):
        html =   "<html><body><font size=+3> Oops,Something has gone wrong!! <br><br><a href='/'>Back<a></font></body></html>"
    if html.__contains__('returncode=0'):
        html =   "<html><body><font size=+3>Script executed successfully!! <br><br><a href='/'>Back<a></body></font></html>"
    return HttpResponse(html)


def showfiles(request):
    arr = os.listdir('automate\\media')
    return render(request, "list-files.html",{"showcity": arr})

def download(request):
    if request.method == 'POST':
        files = request.POST.get('city[]')
        print(files)
        #for i in files:
        #    file1 = i[0]
        #    print(file1)
        file_path = "automate/media/" + str(files)
        if os.path.exists(file_path):
            with open(file_path, 'rb') as fh:
                response = HttpResponse(fh.read(), content_type="application/vnd.ms-excel")
                response['Content-Disposition'] = 'inline; filename=' + os.path.basename(file_path)
            return response
    return HttpResponse()


def upload(request):
    if request.method == "POST":
        images = request.FILES.getlist("images")

        directoryPath = r'C:\Users\mkanniah\automation-v1.1\automate\media'
        # Comparing the returned list to empty list
        if os.listdir(directoryPath) != []:
            print("Some files found in the directory.")
            original = r'C:\Users\mkanniah\automation-v1.1\automate\media'
            target = r'C:\Users\mkanniah\automation-v1.1\automate\media_history'
            file_names = os.listdir(original)
            for file_name in file_names:
                shutil.move(os.path.join(original, file_name), os.path.join(target, file_name))

        for image in images:
            fs = FileSystemStorage()
            if fs.exists(str(image.name)):
                os.remove(os.path.join(settings.MEDIA_ROOT, image.name))
            file_path = fs.save(image.name, image)
            print(image.name)
            messages.success(request, '"' + str(image.name) + '"' + 'File Uploaded Sucessfully..')
    return render(request, 'geniusvoice.html', {'images': image})
