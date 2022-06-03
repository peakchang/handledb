import os, environ
from .base import *


def read_secret(secret_name):
    file = open("/run/secrets/" + secret_name)
    secret = file.read()
    secret = secret.rstrip().lstrip()
    file.close()

    return secret

env = environ.Env(
    # set casting, default value
    DEBUG=(bool, False)
)

# Take environment variables from .env file
environ.Env.read_env(env_file=os.path.join(BASE_DIR, '.env'))

# Quick-start development settings - unsuitable for production
# See https://docs.djangoproject.com/en/4.0/howto/deployment/checklist/
SECRET_KEY = read_secret("DJANGO_SECRET_KEY")

# SECURITY WARNING: don't run with debug turned on in production!
DEBUG = False

ALLOWED_HOSTS = ['*']

# Database
# https://docs.djangoproject.com/en/4.0/ref/settings/#databases

DATABASES = {
    'default': {
        'ENGINE': 'django.db.backends.mysql',
        'NAME': 'handledb',
        'USER': 'pcyadmin',
        'PASSWORD': read_secret("MYSQL_PASSWORD"),
        'HOST': 'mariadb',
        'PORT': '3306',
    }
}
