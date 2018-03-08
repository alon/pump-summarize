#!/usr/bin/env python
from os import system, path
import json
import os

os.system('wget -O pump-summarize.json https://ci.appveyor.com/api/projects/alon/pump-summarize')
with open('pump-summarize.json') as fd:
    js = json.load(fd)
last_job = js['build']['jobs'][0]
job_id = last_job['jobId']
job_updated = last_job['updated']
print('job_id = {}'.format(job_id))
print('job updated: {}'.format(job_updated))
summarize_url = 'https://ci.appveyor.com/api/buildjobs/{}/artifacts/dist%2Fsummarize.zip'.format(job_id)
with open('index.html', 'w') as fd:
    fd.write("""<!DOCTYPE html><html><head><title>Summarize zip file (job_id {}, created {})</title></head><body><a href="{}">summarize.zip</a></body>""".format(
        job_id, job_updated, summarize_url))
if path.exists('summarize.zip'):
    print("removing existing summarize.zip file")
    os.unlink('summarize.zip')
system('wget -O summarize.zip {}'.format(summarize_url))
