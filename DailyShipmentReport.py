import xlwings as xw # need to install
import http.client
import urllib.parse
import json
import datetime
import os
import base64
import pymsgbox   # need to install


def trimHAWB(hawb):

    trim = str(hawb)
    if '.' in trim:
        trim = trim[0: trim.find('.', 0, len(trim))]
    return trim.strip()

def getDHL(hawb):

    d = {}  # d['trackingNumber'] = ','.join(x)
    d['trackingNumber'] = hawb
    d['service'] = 'express'
    params = urllib.parse.urlencode(d)

    headers = {
        'Accept': 'application/json',
        # 'DHL-API-Key': '89tR8kV6AmpAM4UnZN8yHGMNIa8jcKRZ'  # --250
        'DHL-API-Key': 'PIOJQgIXSpxtGvbjnDlkkjQn4ZbzNnBD'  # --1000
    }

    connection = http.client.HTTPSConnection("api-eu.dhl.com")
    connection.request("GET", "/track/shipments?" + params, "", headers)
    response = connection.getresponse()

    data = []
    data.append(response.status)
    data.append(response.reason)
    data.append(json.loads(response.read()))
    connection.close()

    return data

def getEIToken():

    credential = "facebook:ajTubv8eOx4oYoryyJHr"
    encodedBytes = base64.b64encode(credential.encode("utf-8"))
    encodedStr = 'Basic ' + str(encodedBytes, "utf-8")
    params = urllib.parse.urlencode({'grant_type': 'client_credentials'})
    headers = {'Content-Type': 'application/x-www-form-urlencoded', 'Authorization': encodedStr}  # 2000 per hour

    connection = http.client.HTTPSConnection("api.expeditors.com")
    connection.request("POST", "/tracking/v2/oauth2/token?" + params, "", headers)
    response = connection.getresponse()

    data = []
    data.append(response.status)
    data.append(response.reason)
    data.append(json.loads(response.read()))
    connection.close()

    return data

def getEIStatus(hawb, EIToken):

    d = {}
    d['ref'] = hawb
    params = urllib.parse.urlencode(d)

    headers = {}
    headers['Authorization'] = 'Bearer ' + EIToken

    connection = http.client.HTTPSConnection("api.expeditors.com")
    connection.request("GET", "/tracking/v2/shipments/exactmatch?" + params, "", headers)
    response = connection.getresponse()

    data = []
    data.append(response.status)
    data.append(response.reason)
    data.append(json.loads(response.read()))
    connection.close()

    return data

@xw.func(async_mode='threading')
@xw.arg('x', doc='This is HAWB.')
@xw.arg('y', doc= 'This is ''SATUS'', ''TIME'' or ''PICKUP''.')
def getDHLStatus(x, y):

    data = getDHL(trimHAWB(x))
    if data[0] == 200:
        if y.upper() == 'STATUS':
            return data[2]['shipments'][0]['status']['status']
        elif y.upper() == 'TIME':
            return data[2]['shipments'][0]['status']['timestamp']
        elif y.upper() == 'PICKUP':
            for i in data[2]['shipments'][0]['events']:
                if i['description'] == 'Shipment picked up':
                    return i['timestamp']
        else:
            return y + ' is not registered. Check your parameter.'
    else:
        return str(data[0]) + ': ' + str(data[1])


def main():
    dsr_version = '(DSR v1.0)'
    # Validate workbook name
    wb = xw.Book.caller()
    if wb.name != 'DailyShipmentReport.xlsm':
        pymsgbox.alert(text="Please run from DailyShipmentReport.xlsm, did you open that?", title='DSR Message', button='OK')
        exit()
    # Validate Sheet name
    try:
        sht = wb.sheets['Outbound International']
    except:
        pymsgbox.alert(text="Target sheet doesn't exist, did you rename any of your sheet?", title='DSR Message', button='OK')
        exit()
    # Get all columns in the spreadsheet
    headers = sht.range('A2').expand('right')
    for colm in headers:
        if str(colm.value).upper().strip() == 'SHREQ NO.':
            shreq_no_col = colm.column
        elif str(colm.value).upper().strip() == 'SHIPMENT #':
            shipment_no_col = colm.column
        elif str(colm.value).upper().strip() == 'TASK #':
            task_no_col = colm.column
        elif str(colm.value).upper().strip() == 'HAWB #':
            hawb_col = colm.column
        elif str(colm.value).upper().strip() == 'ORIGIN':
            origin_col = colm.column
        elif str(colm.value).upper().strip() == 'DESTINATION':
            destination_col = colm.column
        elif str(colm.value).upper().strip() == 'PROJECT':
            project_col = colm.column
        elif str(colm.value).upper().strip() == 'TITLE':
            title_col = colm.column
        elif str(colm.value).upper().strip() == 'ORDER DROP DATE':
            order_drop_date_col = colm.column
        elif str(colm.value).upper().strip() == 'CI RECEIVED DATE':
            ci_receive_date_col = colm.column
        elif str(colm.value).upper().strip() == 'PICKUP DATE':
            pick_date_col = colm.column
        elif str(colm.value).upper().strip() =='AIRLINE ETA':
            airline_eta_col = colm.column
        elif str(colm.value).upper().strip() == 'LATEST SHIPMENT STATUS':
            latest_shipment_status_col = colm.column
        elif str(colm.value).upper().strip() == 'ACTUAL DELIVERY DATE':
            actual_delivery_date_col = colm.column
        elif str(colm.value).upper().strip() == 'ORIGINAL NBD':
            origianal_nbd_col = colm.column
        elif str(colm.value).upper().strip() == 'REVISED NBD':
            revised_nbd_col = colm.column
        elif str(colm.value).upper().strip() == 'STATUS':
            status_col = colm.column
        elif str(colm.value).upper().strip() == 'REASON CODES':
            reason_code_col = colm.column
        elif str(colm.value).upper().strip() == 'RESPONSIBILITY':
            responsibility_col = colm.column
        elif str(colm.value).upper().strip() == 'COMMENTS':
            comments_col = colm.column
        elif str(colm.value).upper().strip() == 'ACCESS CODE':
            access_code_col = colm.column
        elif str(colm.value).upper().strip() == 'CARRIER':
            carrier_col = colm.column
        elif str(colm.value).upper().strip() == 'BROKER':
            broker_col = colm.column
        elif str(colm.value).upper().strip() == 'WEIGHT AND DIMS':
            weight_and_dims_col = colm.column
        elif str(colm.value).upper().strip() == 'CHARGEABLE WEIGHT (KGS)':
            chargeable_weight_col = colm.column
        elif str(colm.value).upper().strip() == 'PALLET/LOOSE':
            pallet_loose_col = colm.column
        elif str(colm.value).upper().strip() == 'WHITE GLOVE (Y/N)':
            white_glove_col = colm.column
        elif str(colm.value).upper().strip() == 'CURRENCY':
            currency_col = colm.column
        elif str(colm.value).upper().strip() == 'VALUE':
            value_col = colm.column
        elif str(colm.value).upper().strip() == 'DN#':
            DN_col = colm.column
        elif str(colm.value).upper().strip() == 'SO#':
            SO_col = colm.column
        elif str(colm.value).upper().strip() == 'LINE ITEM':
            line_item_col = colm.column
        elif str(colm.value).upper().strip() == 'QUANTITY (PCS)':
            quantity_col = colm.column
        elif str(colm.value).upper().strip() == 'NBD BREACH':
            nbd_breach_col = colm.column
        elif str(colm.value).upper().strip() == 'POTENTIAL BREACH':
            potential_breach_col = colm.column
        elif str(colm.value).upper().strip() == 'DSR UPDATE REMARK':
            dsr_update_remark = colm.column

    res = pymsgbox.confirm(text="Do you want to update open shipment status?", title='DSR Message', buttons=['OK', 'Cancel'])
    if res == 'Cancel':
        exit()
    # Initiate EI token.
    EIToken = getEIToken()
    if EIToken[0] != 200:
        pymsgbox.alert("Cannot get EI token, the response is: " + EIToken[1], 'DSR Message', 'OK')
        exit()

    rng = sht.range('A3').expand('down')

    for i in rng:
        if str(sht.range(i.row, status_col).value).upper().strip() not in ['CLOSE', 'CLOSED', 'CANCEL', 'CANCELLED']:
            if str(sht.range(i.row, hawb_col).value).upper().strip() != '' and sht.range(i.row, hawb_col).value is not None:
                if str(sht.range(i.row, carrier_col).value).upper().strip() == 'DHL':
                    data = getDHL(trimHAWB(sht.range(i.row, hawb_col).value))
                    if data[0] == 200:
                        for e in data[2]['shipments'][0]['events']:
                            if e['description'] == 'Shipment picked up':
                                sht.range(i.row, pick_date_col).value = e['timestamp']
                        sht.range(i.row, latest_shipment_status_col).value = data[2]['shipments'][0]['status']['status']
                        sht.range(i.row, actual_delivery_date_col).value = data[2]['shipments'][0]['status']['timestamp']
                        if str(data[2]['shipments'][0]['status']['status']).upper().strip() == 'DELIVERED':
                            sht.range(i.row, status_col).value = 'CLOSED'
                            sht.range(i.row, dsr_update_remark).value = 'Closed by DSR as shipment is delivered.'
                    else:
                        sht.range(i.row, dsr_update_remark).value = str(data[0]) + ': ' + str(data[1])
                elif str(sht.range(i.row, carrier_col).value).upper().strip() in ['EI', 'EXPEDITORS', 'EI (FB)']:
                    data = getEIStatus(trimHAWB(sht.range(i.row, hawb_col).value), EIToken[2]['access_token'])
                    if data[0] == 200:
                        for e in data[2][0]['events']:
                            if e['description'] == 'Client Called for Pickup':
                                sht.range(i.row, pick_date_col).value = e['time']
                        sht.range(i.row, latest_shipment_status_col).value = data[2][0]['status']
                        sht.range(i.row, actual_delivery_date_col).value = data[2][0]['lastUpdateTime']
                        sht.range(i.row, chargeable_weight_col).value = str(data[2][0]['weight']['value']) + data[2][0]['weight']['units']
                        if str( data[2][0]['status']).upper().strip() == 'SERVICES COMPLETED: DELIVERED':
                            sht.range(i.row, status_col).value = 'CLOSED'
                            sht.range(i.row, dsr_update_remark).value = 'Closed by DSR as shipment is delivered.'
                    else:
                        sht.range(i.row, dsr_update_remark).value = str(data[0]) + ': ' + str(data[1])
                else:
                    sht.range(i.row, dsr_update_remark).value = 'Carrier API is not yet onboarded.'
            else:
                sht.range(i.row, dsr_update_remark).value = 'No HAWB found.'

    sht.range('A1').value = dsr_version + ": Last DSR status run by " + os.getlogin() + " at " + str(datetime.datetime.now())

    res = pymsgbox.confirm(text="DSR status updated, do you want to save?", title='DSR Message', buttons=['OK', 'Cancel'])
    if res == 'OK':
        wb.save()

if __name__ == "__main__":
    xw.books.active.set_mock_caller()
    main()
'''
if __name__ == '__main__':
    xw.serve()
'''