#!/bin/bash
gunicorn --bind=0.0.0.0:8000 --timeout=1800 --workers=2 capacity_api:app
