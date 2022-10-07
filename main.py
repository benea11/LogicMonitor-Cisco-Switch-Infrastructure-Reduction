import logging.config
import hashlib
import base64
import time
from pathlib import Path
import requests
from dotenv import load_dotenv
import os
import logging
import xlsxwriter
import json

"""
Interfaces are not integers! = Peter´s Fix
"""

env_path = Path('.') / '.env'
env_path = os.path.join(env_path)
load_dotenv(env_path)
siteId = #Insert Site Identifier Here
accessId = os.getenv("AccessId")
accessKey = os.getenv("AccessKey")


class LogicMonitor():

    def __init__(self, accessId, accessKey, queryParams, resourcePath, siteId):
        import hmac
        httpVerb = 'GET'
        data = ''
        url = '%%logicmonitor API url' + resourcePath + queryParams
        epoch = str(int(time.time() * 1000))
        requestVars = httpVerb + epoch + data + resourcePath
        hmac = hmac.new(accessKey.encode(), msg=requestVars.encode(), digestmod=hashlib.sha256).hexdigest()
        signature = base64.b64encode(hmac.encode())
        auth = 'LMv1 ' + accessId + ':' + signature.decode() + ':' + epoch
        headers = {'Content-Type': 'application/json', 'Authorization': auth}
        try:
            logger.debug("LogicMonitor API Call")
            response = requests.get(url, data=data, headers=headers)
        except Exception as e:
            logger.info("Cannot connect to the LM API")
            logger.debug(e)
        try:
            i = response.json()
            logger.debug("LogicMonitor API return is OK!")
            logger.debug(i)
        except Exception as e:
            logger.info("LogicMonitor API return was not readable")
            logger.debug(e)
            exit(404)
        self.siteId = siteId
        self.output = i

    def device_List(self):
        get_device_list = self.output["data"]["items"]
        deviceModel = False
        deviceType = False
        device_list = []
        sanitised_device_list = []
        logger.debug("sanitise data (siteID)")
        for device in get_device_list:
            logging.debug(device)
            for property in device["inheritedProperties"]:  # Check if this logic is still required
                if property["name"] == "ctag.siteid":
                    logging.debug("Inputed siteID = " + str(self.siteId))
                    logging.debug("Retrieved SiteID = " + str(property["value"]))
                    if int(self.siteId) == int(property["value"]):  # 21 = 21, 21555 <> 21
                        logger.debug("device matched to siteID, adding to sanitised: " + device["displayName"])
                        sanitised_device_list.append(device)
        for device in sanitised_device_list:
            logger.debug(device)
            for custom_property in device["customProperties"]:
                if custom_property["name"] == "ctag.devicetype":
                    deviceType = custom_property["value"]
            for auto_property in device["autoProperties"]:
                if auto_property["name"] == "auto.endpoint.model":
                    deviceModel = auto_property["value"]
            device_list.append({"id": device["id"],
                                "deviceName": device["displayName"],
                                "deviceModel": deviceModel,
                                "deviceType": deviceType})
            deviceModel = False
            deviceType = False

        try:
            out = [i for i in device_list if ("Switch" in i["deviceType"])]
        except Exception as e:
            logger.critical("Looks like a network device exists in this site that has a incorrect deviceType flag")
            for i in device_list:
                if not isinstance(i["deviceType"], str):
                    logger.critical(i)
            logger.debug(e)
            logger.debug(device_list)
            exit()
        logger.debug(out)
        return out

    def data_sources(self):
        stack_id = False
        id = False
        datasources = self.output["data"]["items"]
        for datasource in datasources:
            if datasource["dataSourceName"] == "Cisco Switch Stack-":
                stack_id = datasource["id"]
                logger.debug("stack ID: " + str(stack_id))
            if datasource["dataSourceName"] == "SNMP_Network_Interfaces":
                id = datasource["id"]
                logger.debug("Data Source ID: " + str(id))
        return stack_id, id

    def data_source_instances(self):
        instances = []
        interfaces = []
        for i in self.output["data"]["items"]:
            logger.debug(i)
            if not i["stopMonitoring"]:
                instances.append(i)
                logger.debug("appended instance")
        for i in instances:
            for j in i["autoProperties"]:
                if j["name"] == "auto.interface.description" and "V" not in j["value"] and "Port" not in j[
                    "value"] and "Tun" not in j["value"] and "ull" not in j["value"]:
                    interfaces.append(j["value"])
                    logger.debug("interface appended" + str(j["value"]))
        return instances, interfaces

    def data_source_instances_stack(self):
        return self.output["data"]["items"]

    def device_parameters(self):
        return self.output["data"]["items"]


def main():
    logger.debug("Site ID: " + str(siteId))
    switch_inventory = []
    # queryParams = '?filter=systemProperties.name:system.staticgroups,systemProperties.value~*Network*-' + str(siteId)

    queryParams = '?filter=inheritedProperties.name:ctag.siteid,inheritedProperties.value:' + str(siteId)
    resourcePath = '/device/devices'
    device_list = LogicMonitor(accessId=accessId, accessKey=accessKey, queryParams=queryParams,
                               resourcePath=resourcePath, siteId=siteId).device_List()

    for device in device_list:
        device_interfaces = execute(device)
        switch_inventory.append(device_interfaces)
        logger.debug("Device Added to Switch Inventory")
    workbook(switch_inventory=switch_inventory)
    return  # Maintain this return


def execute(device):
    deviceId = device["id"]
    resourcePath = '/device/devices/' + str(deviceId) + '/devicedatasources'
    queryParams = '?fields=dataSourceName,id&size=500'

    stack_id, id = LogicMonitor(accessId=accessId, accessKey=accessKey, queryParams=queryParams,
                                resourcePath=resourcePath,
                                siteId=siteId).data_sources()

    resourcePath = '/device/devices/' + str(deviceId) + '/devicedatasources/' + str(id) + '/instances'
    queryParams = ''

    instances, interfaces = LogicMonitor(accessId=accessId, accessKey=accessKey, queryParams=queryParams,
                                         resourcePath=resourcePath,
                                         siteId=siteId).data_source_instances()

    resourcePath = '/device/devices/' + str(deviceId) + '/devicedatasources/' + str(stack_id) + '/instances'
    queryParams = ''

    get_datasource_instances_stack = LogicMonitor(accessId=accessId, accessKey=accessKey, queryParams=queryParams,
                                                  resourcePath=resourcePath,
                                                  siteId=siteId).data_source_instances_stack()

    stack_members = len([ele for ele in get_datasource_instances_stack if isinstance(ele, dict)])
    device["stacks"] = stack_members
    logger.debug("stack members: " + str(stack_members))
    logger.debug("interfaces: " + str(interfaces))
    output = interface_realiser(interface_list=interfaces, devicedict=device)
    return output


def interface_realiser(interface_list, devicedict):
    output = {}
    SW1 = []
    SW2 = []
    SW3 = []
    SW4 = []
    logger.debug("device dict: " + str(devicedict))
    logger.debug("interace_list: " + str(interface_list))
    for interface in interface_list:
        slash_count = interface.count("/")
        break
    for interface in interface_list:
        split = interface.split("/")
        logger.debug("The first interface split: " + str(split))
        # FIX THIS BUG PROPERLY..
        try:
            switch = int(split[0].split("net")[1])
            if switch == 0:
                switch = 1
            logger.debug("the switch ID is: " + str(switch))
        except Exception as e:
            logger.debug(split)
            logger.debug(e)
            continue

        if switch == 1:
            SW1.append(str(split[slash_count]))
        elif switch == 2:
            SW2.append(split[slash_count])
        elif switch == 3:
            SW3.append(split[slash_count])
        elif switch == 4:
            SW4.append(split[slash_count])
    try:  # Check this
        port_count = str(devicedict["deviceModel"]).split("-")
        if "C9" in port_count[0]:
            port_count = str(port_count[1])
            numeric_filter = filter(str.isdigit, port_count)
            port_count = "".join(numeric_filter)
        else:
            port_count = str(port_count[2])
            numeric_filter = filter(str.isdigit, port_count)
            port_count = "".join(numeric_filter)
    except:
        port_count = False

    output["name"] = devicedict["deviceName"]
    output["model"] = devicedict["deviceModel"]
    output["portCount"] = port_count
    output["stacks"] = devicedict["stacks"]

    output["interfaces"] = {}

    switch1 = [i for j, i in enumerate(SW1) if i not in SW1[
                                                        :j]]  # Remove duplicates (when uplink and access interfaces exist with the same interface number
    output["interfaces"].update({"Stack Member 1": switch1})
    output["SW1.Occupied"] = len(switch1)
    if output["stacks"] > 1:
        switch2 = [i for j, i in enumerate(SW2) if i not in SW2[
                                                            :j]]  # Remove duplicates (when uplink and access interfaces exist with the same interface number
        output["interfaces"].update({"Stack Member 2": switch2})
        output["SW2.Occupied"] = len(switch2)
    if output["stacks"] > 2:
        switch3 = [i for j, i in enumerate(SW3) if i not in SW3[
                                                            :j]]  # Remove duplicates (when uplink and access interfaces exist with the same interface number
        output["interfaces"].update({"Stack Member 3": switch3})
        output["SW3.Occupied"] = len(switch3)
    if output["stacks"] > 3:
        switch4 = [i for j, i in enumerate(SW4) if i not in SW4[
                                                            :j]]  # Remove duplicates (when uplink and access interfaces exist with the same interface number
        output["interfaces"].update({"Stack Member 4": switch4})
        output["SW4.Occupied"] = len(switch4)

    logger.warning(output)
    return output


def workbook(switch_inventory):
    f = open("port_to_y_coordinate.json")
    port_to_coordinate_y = json.load(f)
    f = open("port_to_x_coordinate.json")
    port_to_coordinate_x = json.load(f)
    f = open("lm_model_to_port_capacity.json")
    lm_model_to_port_capacity = json.load(f)

    row_count = 1
    workbook = xlsxwriter.Workbook(str(siteId) + ".xlsx")
    worksheet = workbook.add_worksheet("Switchport Capacity")
    worksheet.set_column("A:A", 15)
    worksheet.set_column("B:AW", 5)
    cell_format_used = workbook.add_format()
    cell_format_used.set_font_color('white')
    cell_format_used.set_bg_color('red')
    cell_format_used.set_border(1)
    cell_format_free = workbook.add_format()
    cell_format_free.set_font_color('white')
    cell_format_free.set_bg_color('green')
    cell_format_free.set_border(1)
    percentage_format = workbook.add_format({'num_format': "0%"})
    switch_name_format = workbook.add_format()
    switch_name_format.set_bold()
    stack_member_format = workbook.add_format()
    stack_member_format.set_italic()
    stack_member_format.set_align("right")

    for switch in switch_inventory:
        logger.debug("Row Count" + str(row_count))
        worksheet.write(0 + row_count, 0, switch["name"], switch_name_format)
        for stack_member in switch["interfaces"]:
            worksheet.write(1 + row_count, 0, stack_member, stack_member_format)
            try:
                used_interfaces = [int(x) for x in switch["interfaces"][stack_member]]
            except Exception as e:
                logger.warning(switch["name"] + " Interfaces are not integers!")
                logger.debug(e)
                break
            logger.debug(switch["model"])
            if not switch["model"]:  # Force to 48 port switch
                logger.debug(switch["name"] + ": unknown port capacity for the switch model " + str(
                    switch["model"]) + " setting to 48")

            try:
                max_interface = lm_model_to_port_capacity[switch["model"]]
            except Exception as e:
                logger.debug(e)
                logger.info("unable to determine total amount of interfaces, setting 48 ports: " + str(switch["name"]))
                max_interface = 48

            if max_interface == 6:
                complete_interfaces = range(1, 6)
            elif max_interface == 8:
                complete_interfaces = range(1, 8)
            elif max_interface == 12:
                complete_interfaces = range(1, 12)
            elif max_interface == 16:
                complete_interfaces = range(1, 16)
            elif max_interface == 24:
                complete_interfaces = range(1, 24)
            elif max_interface == 32:
                complete_interfaces = range(1, 32)
            elif max_interface == 40:
                complete_interfaces = range(1, 40)
            elif max_interface == 48:
                complete_interfaces = range(1, 48)

            unused_interfaces = list(sorted(set(complete_interfaces) - set(used_interfaces)))
            for interface in unused_interfaces:
                try:
                    logger.debug("GREEN" + switch["name"] + " interface " + str(interface) + " " + str(
                        port_to_coordinate_y[str(interface)]) + " " + str(port_to_coordinate_x[str(interface)]))
                    worksheet.write(int(port_to_coordinate_y[str(interface)]) + row_count,
                                    int(port_to_coordinate_x[str(interface)]), str(interface), cell_format_free)
                except Exception as e:
                    logger.debug(e)
                    logger.warning(switch["name"] + " unable to parse the interfaces: " + str(unused_interfaces))
                    break
            for interface in switch["interfaces"][stack_member]:
                try:
                    if int(interface) < int(max_interface)+1:
                        logger.debug("RED" + switch["name"] + " interface " + str(interface) + " " + str(
                            port_to_coordinate_y[str(interface)]) + " " + str(port_to_coordinate_x[str(interface)]))
                        worksheet.write(int(port_to_coordinate_y[interface]) + row_count,
                                        int(port_to_coordinate_x[interface]), interface, cell_format_used)
                    else:
                        logger.debug(str(interface) + " exceeds the expected max interfaces, skipping.")
                except Exception as e:
                    logger.debug(e)
                    logger.warning(switch["name"] + " unable to parse the interfaces: " + str(switch))
                    break

            try:
                capacity = len(switch["interfaces"][stack_member]) / max_interface
                worksheet.write(2 + row_count, 0, capacity, percentage_format)
            except Exception as e:
                logger.critical("Unable to calculate port capacity for " + switch["name"] +
                                " Device Model " + switch["model"])
                logger.debug(e)


            row_count += 3
        row_count += 2
    workbook.close()


if __name__ == "__main__":
    logging.config.fileConfig(fname='logging.conf', disable_existing_loggers=True)
    stream = logging.StreamHandler()
    streamformat = logging.Formatter("%(levelname)s:%(module)s:%(lineno)d:%(message)s")
    stream.setFormatter(streamformat)
    logger = logging.getLogger(__name__)
    logger.propagate = False
    logger.addHandler(stream)
    main()
