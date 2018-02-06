#!/usr/bin/env python
from requests import get
from os import system, path
import os


response = get('https://ci.appveyor.com/api/projects/alon/pump-summarize')
js = response.json()
last_job = js['build']['jobs'][0]
job_id = ['jobId']
print('job_id = {}'.format(job_id))
print('job updated: {}'.format(last_job['updated']))
summarize_url = 'https://ci.appveyor.com/api/buildjobs/{}/artifacts/dist%2Fsummarize.zip'.format(job_id)
with open('index.html', 'w') as fd:
    fd.write("""<!DOCTYPE html><html><head><title>Summarize zip file</title></head><body><a href="{}">summarize.zip</a></body>""".format(summarize_url))
if path.exists('summarize.zip'):
    print("removing existing summarize.zip file")
    os.unlink('summarize.zip')
system('wget -O summarize.zip {}'.format(summarize_url))
