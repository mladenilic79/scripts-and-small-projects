
import os
import shutil
from time import sleep

reports_dir = "reports"
reports_dir_single_customers = reports_dir + os.sep + "single_customers_reports"

source_dir = "source"
source_dir_audits = source_dir + os.sep + "audits"
source_dir_imacs = source_dir + os.sep + "imacs"
source_dir_agents = source_dir + os.sep + "agents"
source_dir_devices = source_dir + os.sep + "devices"
source_dir_inventories = source_dir + os.sep + "inventories"

try:
    shutil.rmtree(reports_dir)
except:
    print("error deleting reports directory structure")
    sleep(3)
# wait for windows to finish
sleep(3)
os.mkdir(reports_dir)
os.mkdir(reports_dir_single_customers)

try:
    shutil.rmtree(source_dir)
except:
    print("error deleting source directory structure")
    sleep(3)
# wait for windows to finish
sleep(3)
os.mkdir(source_dir)
os.mkdir(source_dir_audits)
os.mkdir(source_dir_imacs)
os.mkdir(source_dir_agents)
os.mkdir(source_dir_devices)
os.mkdir(source_dir_inventories)

print("configuration ready")
sleep(3)
