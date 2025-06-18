#!/usr/bin/env bash
# exit on error
set -o errexit

pip install -r requirements.txt

cd tiketon
mkdir -p staticfiles
python manage.py collectstatic --no-input --clear
python manage.py migrate 