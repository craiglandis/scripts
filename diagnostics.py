#!/usr/bin/python

import re
import json
import subprocess
import argparse
import sys

try:
  from termcolor import colored
except ImportError, e:
  print "termcolor module is required"
try:
  from azure.servicebus import ServiceBusService
except e:
  print "azure module is required"

# == Data Collection ==

def fstab():
  fields = ['fs_spec','fs_file','fs_vfstype','fs_mntops','fs_type','fs_freq','fs_passno']
  f = open('/etc/fstab', 'r')
  lines = f.readlines()
  f.close()
  data = map(lambda x: x.strip(), lines)
  data = filter(lambda x: not x.startswith('#') and x != '', data)
  data = map(lambda x: re.split('\s+',x), data)
  data = map(lambda x: dict(zip(fields,x)), data)
  return data

def diskstats():
  fields=['major','minor','name','reads','reads_merged','sectors','read_time','writes','writes_merged','sectors2','time','ios','io_time','weighted_time']
  int_fields = list(fields)
  int_fields.remove('name')
  f = open('/proc/diskstats','r')
  lines = f.readlines()
  f.close()
  data = map(lambda x: x.strip(),lines)
  data = map(lambda x: re.split('\s+',x),data)
  data = map(lambda x: dict(zip(fields, x)),data)
  data = map(lambda x: dict_filtered_map(int,int_fields,x),data)
  return data

def loadavg():
  fields = ['1','5','15','pid','p_running','p_total']
  f = open('/proc/loadavg','r')
  contents = f.read().strip()
  f.close
  data = re.split('\s+',contents)
  data.extend(data[3].split('/'))
  data.pop(3)
  data = dict(zip(fields,data))
  data = dict_filtered_map(float,['1','5','15'],data)
  data = dict_filtered_map(int,['pid','p_running','p_total'],data)
  return data

def meminfo():
  f = open('/proc/meminfo')
  lines = f.readlines()
  f.close()
  lines = map(lambda x: re.split(':',x), lines)
  data = dict(map(lambda x: (x[0].strip(),x[1].strip()), lines))
  return data

def vmstat():
  f = open('/proc/vmstat')
  lines = f.readlines()
  f.close()
  data = map(lambda x: re.split('\s+',x), lines)
  data = dict(map(lambda x: (x[0],int(x[1])), data))
  return data

def procsysfs():
  data = dict()
  f = open('/proc/sys/fs/file-max')
  data['file-max']=int(f.read().strip())
  f.close()
  f = open('/proc/sys/fs/inode-nr')
  contents = f.read().strip()
  f.close()
  r=re.split('\s+', contents)
  data['inodes']=r[0]
  data['inodes-used']=r[1]
  return data

def inodes_by_mp():
  headers=['filesystem','inodes','inodes_used','inodes_free','inodes_used_percent','mountpoint']
  data = dict()
  p = subprocess.Popen(['df','-i'], stdout=subprocess.PIPE)
  out,err = p.communicate()
  lines = out.strip().split('\n')
  lines = lines[1:]
  data = map(lambda x: dict(zip(headers,re.split('\s+',x))), lines)
  data = map(lambda x: dict_filtered_map(int,['inodes','inodes_used','inodes_free'],x), data)
  return data

def waagent(): 
  f = open('/etc/waagent.conf')
  lines = f.readlines()
  f.close()
  data = map(lambda x: x.split('#')[0].strip(), lines)
  data = filter(lambda x: not x=='', data)
  data = map(lambda x: x.split('='), data)
  data = dict(map (lambda x: (x[0],x[1]), data))
  return data

# == Helper Functions ==
def report(name,level,category,description,link):
  return {'name': name, 'level': level, 'category': category, 'description': description, 'link':link}

def dict_filtered_map(f,attributes,obj):
  g = lambda x,y: f(y) if x in attributes else y
  return {k: g(k,v) for k,v in obj.items()}

def human_readable(obj):
  return colored(obj['name'], obj['level']) + ':    ' + obj['description'] + '\n'

# == Diagnostics ==
class Diagnostics:
  def __init__(self,telemetry,case):
    self.sbkey = '5QcCwDvDXEcrvJrQs/upAi+amTRMXSlGNtIztknUnAA='
    self.sbs = ServiceBusService('linuxdiagnostics-ns', shared_access_key_name='PublicProducer', shared_access_key_value=self.sbkey)
    self.telemetry = telemetry
    self.case = case
    self.report = dict()
    self.data = dict()
    try:
      self.data['fstab'] = fstab()
      self.data['vmstat'] = vmstat()
      self.data['diskstats'] = diskstats()
      self.data['meminfo'] = meminfo()
      self.data['loadavg'] = loadavg()
      self.data['proc.sys.fs'] = procsysfs()
      self.data['inodes_by_mount'] = inodes_by_mp()
      self.data['waagent'] = waagent()
      self.check_fstab_uuid()
      self.check_resourcedisk_inodes()
      self.success_telemetry()
    except:
      self.error_telemetry(str(sys.last_value))
      exit()

  def print_report(self,human):
    if(human):
      for item in map(lambda x: human_readable(self.report[x]),self.report):
        print item
    else:
      print json.dumps(self.report)

# == Telemetry ==
  def success_telemetry(self):
    if(self.telemetry):
      event = {'event': 'success', 'report': self.report, 'data': self.data}
      if(self.case):
        event['case'] = self.case
      self.sbs.send_event('linuxdiagnostics',json.dumps(event))

  def error_telemetry(self,e):
    if(self.telemetry):
      event = {'event': 'error', 'error': e, 'report': self.report, 'data':self.data}
      if(self.case):
        event['case'] = self.case
      sbs.send_event('linuxdiagnostics',json.dumps(event))

# == Diagnostics ==
  def check_fstab_uuid(self):
    uuids = map (lambda r:bool(re.match('UUID',r['fs_spec'],re.I)), self.data['fstab'])
    uuids = filter(lambda x: not x, uuids)
    if(len(uuids) > 0):
      self.report['linux.uuid']= report("fstab doesn't use UUIDs", 'red', 'best_practices', "/etc/fstab isn't using UUIDs to reference its volumes.", '')

  def check_resourcedisk_inodes(self):
    mountpoint = self.data['waagent']['ResourceDisk.MountPoint']
    x=filter(lambda x: x['mountpoint']==mountpoint ,self.data['inodes_by_mount'])[0]
    if(x['inodes_used'] > 20):
      self.report['azure.resource_disk_usage'] = report('Resource Disk Usage', 'yellow', 'data_loss',"This instance appears to be using the resource disk. Data stored on this mountpoint may be lost.", '')

def main():
  parser = argparse.ArgumentParser(description='Run automated diagnostics')
  parser.add_argument('-y', action='store_true', help='Skip prompts')
  parser.add_argument('--anonymous', action='store_true', help='No telemetry')
  parser.add_argument('--case', help='Link output with a case, for use with Azure Support')
  parser.add_argument('--human', action='store_true', help='Write human-readable text to stdout')
  parser.add_argument('--feedback', action='store_true', help='Provide suggestions and/or feedback')

  args=parser.parse_args()

  if(args.feedback):
    suggestion = raw_input("Suggestion (Press return to submit):  ")
    sbkey = '5QcCwDvDXEcrvJrQs/upAi+amTRMXSlGNtIztknUnAA='
    sbs = ServiceBusService('linuxdiagnostics-ns', shared_access_key_name='PublicProducer', shared_access_key_value=sbkey)
    event = {'event': 'feedback', 'feedback': suggestion}
    sbs.send_event('linuxdiagnostics',json.dumps(event))
    exit()

  if(args.anonymous):
    if(args.case == True):
      print "Cannot report case diagnostics in anonymous mode."
    else:
      Diagnostics(False,args.case).print_report(args.human)
  else:
    if(args.y):
      Diagnostics(True,args.case).print_report(args.human)
    else:
      prompt="In order to streamline support and drive continuous improvement, this script may transmit system data to Microsoft.\n Do you consent to this? (y/n):"
      consent=raw_input(prompt)
      if(consent=='y' or consent=='Y'):
        Diagnostics(True,args.case).print_report(args.human)
      else:
        print("To run without telemetry, use the --anonymous flag")

if __name__ == "__main__":
  main()
