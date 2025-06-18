#!/usr/bin/env bash
# exit on error
set -o errexit

pip install -r requirements.txt

python tiketon/manage.py collectstatic --no-input
python tiketon/manage.py migrate 