import time
from django.shortcuts import render, redirect
import subprocess
import os
# Create your views here.
def indexpage(request):
    return render(request,"index.html")

def my(request):
    print("main")
    subprocess.run(['python', 'main.py'], shell=True)  # in list first is command  and second is argument
    return redirect(my)