#!/bin/bash

echo "--- Starting Briefprogramm ---"

# Get name of Wifi device
mywifidevice=$(nmcli device | grep wifi\ | awk '{print $1}')

# Get name of Wifi SSID
myssid=$(nmcli -t -f name,device connection show --active | grep $mywifidevice | cut -d\: -f1)

echo "Wifi SSID:" $myssid

# Activate virtual environment
source venv/bin/activate

# Start Briefprogramm.py with wifi ssid as argument
echo -e "\n-- Briefprogramm.py --"
python3 Briefprogramm.py $myssid

