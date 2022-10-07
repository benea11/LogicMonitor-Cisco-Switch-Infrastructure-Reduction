<div id="top"></div>
<br />
<div align="center">
  <h3 align="center">Infrastructure Reduction Realiser</h3>

  <p align="center">
    Calculate switchport usage via Logic Monitor and produce a excel report.
    <br />
   </p>
</div>

## Getting Started
### Prerequisites

.env file with the Logic Monitor AccessId and AccessKey.

### Installation
1. Install pre-requisites in the requirements.txt file (use a VENV):
  ```
  pip3 -r requirements.txt
  ```
2. Modify the Site Identifier variable, line 16, and input the logicmonitor API URL
```
siteId = xxxxxx
  ```
3. Run the script
```
python3 main.py
  ```
