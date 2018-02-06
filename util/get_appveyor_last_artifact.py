#!/usr/bin/env python
from requests import get
from os import system


response = get('https://ci.appveyor.com/api/projects/alon/pump-summarize')
js = response.json()
job_id = js['build']['jobs'][0]['jobId']
print('job_id = {}'.format(job_id))
summarize_url = 'https://ci.appveyor.com/api/buildjobs/{}/artifacts/dist%2Fsummarize.zip'.format(job_id)
with open('index.html', 'w') as fd:
    fd.write("""<!DOCTYPE html><html><head><title>Summarize zip file</title></head><body><a href="{}">summarize.zip</a></body>""".format(summarize_url))
system('wget -q -c -O summarize.zip {}'.format(summarize_url))
