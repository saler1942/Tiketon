#!/usr/bin/env bash
# exit on error
set -e

pip install -r requirements.txt
mkdir -p staticfiles
python manage.py collectstatic --no-input --clear
python manage.py migrate 