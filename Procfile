web: gunicorn spm_live.wsgi --log-file - 
#or works good with external database
web: python manage.py migrate && gunicorn spm_live.wsgi