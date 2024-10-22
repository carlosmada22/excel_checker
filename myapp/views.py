from django.shortcuts import render, redirect
from django.http import JsonResponse, HttpResponse
from django.contrib import messages
from pybis import Openbis
from myapp.utils import name_checker, content_checker, entity_checker, generate_csv_and_download
import logging

# Get an instance of the logger for the app (replace 'myapp' with your app name)
logger = logging.getLogger('myapp')

def homepage(request):
    context = {}
    if request.method == "POST":
        if "upload" in request.POST and request.FILES.get("file"):
            uploaded_file = request.FILES["file"]
            # Check file extension
            if uploaded_file.name.endswith(('.xls', '.xlsx')):
                try:
                    
                    url = f"https://devel.datastore.bam.de/"
                    o = Openbis(url)
                    o.login("usr", "psswd", save_token=True)
                    
                    file_name = uploaded_file.name

                    result_name, code, name_ok = name_checker(file_name)
                    result_content = str(content_checker(uploaded_file, name_ok))
                    result_entity = str(entity_checker(uploaded_file, o))
                    logger.info(f"Type {type(file_name)} of file {file_name}")
                    result_format = "CHECKED NAME:" + "\n----------------------------\n" + result_name + "\n" + "\nCHECKED CONTENT:" + "\n----------------------------\n" + result_content + "\n" + "\nCHECKED ENTITY" + "\n----------------------------\n" + result_entity


                    context["result"] = result_format
                    context["file_name"] = file_name
                    context["code"] = code

                except Exception as e:
                    context["error"] = f"Error processing file: {str(e)}"
            else:
                context["error"] = "Invalid file type. Only .xls and .xlsx files are allowed."

    return render(request, 'homepage.html', context)

# View to handle instance check and CSV generation
def check_instance(request):
    if request.method == 'POST':
        instance = request.POST.get('instance')

        # Simulate fetching data from the OpenBIS instance
        url = f"https://devel.datastore.bam.de/"
        o = Openbis(url)
        o.login("usr", "psswd", save_token=True)

        # Generate CSV data and capture the rows being written
        csv_rows, csv_file, masterdata = generate_csv_and_download(o, instance)

        # Store the CSV content in the session for later download
        request.session[instance] = csv_file
        request.session[f'{instance}_masterdata'] = masterdata

        # Return the rows back to the client for display
        return JsonResponse({
            'rows': csv_rows,  # This is JSON serializable (a list of lists)
            'csv_file': instance,  # The filename is just the instance name
            'masterdata': masterdata
        })

    return JsonResponse({'error': 'Invalid request'}, status=400)

# View to handle the CSV file download (in memory)
def download_csv(request, filename):
    csv_file = request.session.get(filename)  # Retrieve the in-memory CSV file from the session
    if csv_file:
        response = HttpResponse(csv_file, content_type='text/csv')
        response['Content-Disposition'] = f'attachment; filename="{filename}.csv"'
        return response
    else:
        return HttpResponse('File not found', status=404)
    

def login(request):
    if request.method == 'POST':
        username = request.POST.get('username')
        password = request.POST.get('password')

        # Simulate OpenBIS login
        try:
            url = f"https://devel.datastore.bam.de/"
            o = Openbis(url)
            o.login(username, password, save_token=True)

            # If login is successful, save credentials to the session
            request.session['openbis_username'] = username
            request.session['openbis_password'] = password

            # Redirect to homepage after successful login
            return redirect('homepage')

        except Exception as e:
            # Show error if login fails
            return render(request, 'login.html', {
                'error': 'Login failed. Please check your username and password.'
            })
    
    # If GET request, show login page
    return render(request, 'login.html')