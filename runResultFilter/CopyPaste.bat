@echo off
REM Map the network drive (replace \\server\share with your UNC path, and Z: with your preferred drive letter)
net use U:\\gar.corp.intel.com\ec\proj\mdl\pg\intel\engineering\dev\team_client_cmv\users\Hs\script\Process-Improvement\runResultFilter /persistent:yes

REM Run the Python script (replace U:\folder\your_script.py with the full path on the mapped drive)
python U:\users\Hs\script\Process-Improvement\runResultFilter\filterfx_pilot_release.py