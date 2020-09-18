import json
import xlsxwriter
import os
from airtable import Airtable
from io import BytesIO
from datetime import datetime
import urllib

def read_field(obj, *fields):
    try:
        for field in fields:
            if isinstance(obj[field], list):
                obj = obj[field][0]
            else:
                obj = obj[field]
        return obj
    except:
        return ''

def get_domestic_shipments_from_airtable(app_id, secret_key, shipment_group_id):
    try:
        # initialize airtable tables
        tbl_domestic_shipments = Airtable(app_id, 'Domestic Shipments', secret_key)
        tbl_fclist = Airtable(app_id, 'FCList', secret_key)
        tbl_domestic_shipment_line_item = Airtable(app_id, 'DomesticShipmentLineItem', secret_key)
        tbl_skus = Airtable(app_id, 'SKUS', secret_key)
        tbl_packaging_profile = Airtable(app_id, 'PackagingProfile', secret_key)
        tbl_shipment_group = Airtable(app_id, 'ShipmentGroup', secret_key)

        print('##### Getting data from Airtable started #####')
        
        shipment_group = tbl_shipment_group.get(shipment_group_id)
        
        # get all domestic shipments
        domestic_shipments = []
        
        for domestic_shipment_id in shipment_group['fields']['DomesticShipments']:
            domestic_shipment = tbl_domestic_shipments.get(domestic_shipment_id)
            
            if 'Cosignee Name' in shipment_group['fields']:
                domestic_shipment['cosignee'] = shipment_group['fields']['Cosignee Name']
            
            # get shipment information
            if 'FCID' in domestic_shipment['fields']:
                domestic_shipment['shipment'] = tbl_fclist.get(domestic_shipment['fields']['FCID'][0])
            
            # get line items
            line_items = []
            for domestic_shipment_line_item in domestic_shipment['fields']['LineItems']:
                line_item = tbl_domestic_shipment_line_item.get(domestic_shipment_line_item)
                line_item['sku'] = tbl_skus.get(line_item['fields']['SKU'][0])
                line_item['packaging_profile'] = tbl_packaging_profile.get(line_item['fields']['PackagingProfile'][0])
                line_items.append(line_item)
            domestic_shipment['line_items'] = line_items

            domestic_shipments.append(domestic_shipment)
        print('##### Getting data from Airtable finished #####')
        return domestic_shipments
    except Exception as e:
        print('Error getting domestic shipments from Airtable: ' + str(e))
        raise ValueError('Error getting domestic shipments from Airtable: ' + str(e))


def upload_packaging_list_to_airtable(app_id, secret_key, record_id, list_url):
    try:
        # initialize airtable tables
        print('##### Uploading packaging list to Airtable started #####')
        tbl_shipment_group = Airtable(app_id, 'ShipmentGroup', secret_key)
        shipment_group_record = tbl_shipment_group.get(record_id)
        packing_lists = []
        if 'PackingLists Generated' in shipment_group_record['fields']:
            packing_lists = shipment_group_record['fields']['PackingLists Generated']
        packing_lists.append({'url': list_url})
        tbl_shipment_group.update(shipment_group_record['id'], {
            'PackingLists Generated': packing_lists
        })
        print('##### Uploading packaging list to Airtable finished #####')
    except Exception as e:
        print('Error uploading packaging list to Airtable: ' + str(e))
        raise ValueError('Error uploading packaging list to Airtable: ' + str(e))

def get_skus(domestic_shipments):
    skus = []
    for domestic_shipment in domestic_shipments:
        for line_item in domestic_shipment['line_items']:
            sku = line_item['sku']['fields']['SKU']
            if not sku in skus:
                skus.append(sku)
    return skus


def generate_excel_file(domestic_shipments, file_name=None):
    try: 
        # Create an new Excel file and add a worksheet.
        print('##### Generating packaging list started #####')
        output = BytesIO()
        if file_name:
            workbook = xlsxwriter.Workbook(file_name)
        else:
            workbook = xlsxwriter.Workbook(output)
        
        worksheet = workbook.add_worksheet()

        # apply styles
        worksheet.set_column('A:A', 15)
        worksheet.set_column('B:B', 15)
        worksheet.set_column('C:C', 13)
        worksheet.set_column('D:D', 13)
        worksheet.set_column('E:E', 15)
        worksheet.set_column('F:F', 13)
        worksheet.set_column('G:G', 15)
        worksheet.set_column('L:L', 14)
        worksheet.set_column('M:M', 20)
        worksheet.set_row(0, 20)
        worksheet.set_row(2, 50)

        # formats
        highlight_format = workbook.add_format({
            'fg_color': 'yellow'
        })
        rect_format = workbook.add_format({
            'border': 1,
            'align': 'center'
        })
        table_header_format = workbook.add_format({
            'fg_color': '#FCE4D6',
            'border': 1,
            'valign': 'bottom',
            'text_wrap': True
        })
        table_header_without_border_format = workbook.add_format({
            'fg_color': '#FCE4D6',
            'valign': 'bottom'
        })
        box_header_format = workbook.add_format({
            'fg_color': 'red',
            'border': 1,
            'valign': 'bottom'
        })
        rect_box_format = workbook.add_format({
            'fg_color': 'red',
            'border': 1,
            'align': 'center'
        })
        number_format = workbook.add_format({
            'num_format': '0.00'
        })
        rect_number_format = workbook.add_format({
            'num_format': '0.00',
            'border': 1,
            'align': 'center'
        })
        rect_integer_format = workbook.add_format({
            'num_format': '0',
            'border': 1,
            'align': 'center'
        })
        highlight_number_format = workbook.add_format({
            'fg_color': 'yellow',
            'num_format': '0.00'
        })
        text_wrap_format = workbook.add_format({
            'text_wrap': True
        })
        border_top_format = workbook.add_format({
            'top': 1
        })
        border_thick_top_format = workbook.add_format({
            'top': 2
        })
        integer_format = workbook.add_format({
            'num_format': '0'
        })
        
        # fill in basic information
        worksheet.merge_range('A1:M1', 'Packing List', workbook.add_format({
            'bold': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 14
        }))
        worksheet.merge_range('A2:C2', 'Shipment Summary', workbook.add_format({
            'bold': 1,
            'align': 'center',
            'valign': 'vcenter'
        }))

        worksheet.write('B3', 'Units Shipped 订货数量（套）', text_wrap_format)
        worksheet.write('C3', 'Number of Cases/箱数量', text_wrap_format)

        worksheet.write('I4', 'Ship To', workbook.add_format({
            'fg_color': 'yellow',
            'bold': 1,
            'left': 2,
            'top': 2,
            'bottom': 2
        }))
        worksheet.write('J4', read_field(domestic_shipments[0], 'cosignee'), workbook.add_format({
            'fg_color': 'yellow',
            'top': 2,
            'bottom': 2
        }))
        workbook.define_name('VBA_ShipTo', '=Sheet1!$J$4')
        worksheet.write('K4', '', workbook.add_format({
            'fg_color': 'yellow',
            'top': 2,
            'bottom': 2
        }))
        worksheet.write('L4', '', workbook.add_format({
            'fg_color': 'yellow',
            'right': 2,
            'top': 2,
            'bottom': 2
        }))

        skus = get_skus(domestic_shipments)
        domestic_shipment_line = 7 + len(skus)
        # loop through all domestic shipments
        for domestic_shipment in domestic_shipments:
            # draw top thick border
            for i in range(12):
                worksheet.write(domestic_shipment_line - 1, i, '', border_thick_top_format)

            # fill in shipment information
            shipment = read_field(domestic_shipment, 'shipment')
            domestic_shipment_line_start = domestic_shipment_line + 10
            domestic_shipment_line_end = domestic_shipment_line_start + len(domestic_shipment['line_items'])

            # Fulfillment Center
            worksheet.write(domestic_shipment_line, 0, 'Fulfillment Center')
            worksheet.write(domestic_shipment_line, 1, read_field(shipment, 'fields', 'FCID'), highlight_format)
            worksheet.write(domestic_shipment_line, 2, '', highlight_format)

            # Shipment ID
            worksheet.write(domestic_shipment_line + 1, 0, 'Shipment ID')
            worksheet.write(domestic_shipment_line + 1, 1, read_field(domestic_shipment, 'fields', 'FBA Shipment ID'), highlight_format)
            worksheet.write(domestic_shipment_line + 1, 2, '', highlight_format)

            # Reference ID
            worksheet.write(domestic_shipment_line + 2, 0, 'Reference ID')
            worksheet.write(domestic_shipment_line + 2, 1, read_field(domestic_shipment, 'fields', 'AMZReferenceID'), highlight_format)
            worksheet.write(domestic_shipment_line + 2, 2, '', highlight_format)

            # Ship To
            worksheet.write(domestic_shipment_line + 3, 0, 'Ship to')
            worksheet.write_formula(domestic_shipment_line + 3, 1, '=VBA_ShipTo', highlight_format)
            worksheet.write(domestic_shipment_line + 3, 2, '', highlight_format)

            # Total KG
            worksheet.write(domestic_shipment_line + 3, 5, 'Total KG 总公斤')
            worksheet.write_formula(domestic_shipment_line + 3, 8, '=SUM($F${}:$F${})'.format(domestic_shipment_line_start + 1, domestic_shipment_line_end + 1), number_format)
            
            # Total CBM
            worksheet.write(domestic_shipment_line + 4, 5, 'Total CBM 总立方米')
            worksheet.write_formula(domestic_shipment_line + 4, 8, '=SUM($G${}:$G${})'.format(domestic_shipment_line_start + 1, domestic_shipment_line_end + 1), number_format)

            # Number of Cases
            worksheet.write(domestic_shipment_line + 5, 5, 'Number of Cases/箱数量')
            worksheet.write_formula(domestic_shipment_line + 5, 8, '=SUM($D${}:$D${})'.format(domestic_shipment_line_start + 1, domestic_shipment_line_end + 1), integer_format)

            # Units Shipped
            worksheet.write(domestic_shipment_line + 6, 5, 'Units Shipped 订货数量（套）')
            worksheet.write_formula(domestic_shipment_line + 6, 8, '=SUM($E${}:$E${})'.format(domestic_shipment_line_start + 1, domestic_shipment_line_end + 1), integer_format)

            # Cosignee
            worksheet.write(domestic_shipment_line + 4, 0, 'Cosignee')
            worksheet.write(domestic_shipment_line + 4, 1, read_field(domestic_shipment, 'cosignee'), highlight_format)
            worksheet.write(domestic_shipment_line + 4, 2, '', highlight_format)
            address = read_field(shipment, 'fields', 'FCAddress').split(', ', 1)
            worksheet.write(domestic_shipment_line + 5, 1, address[0], highlight_format)
            worksheet.write(domestic_shipment_line + 5, 2, '', highlight_format)
            try:
                worksheet.write(domestic_shipment_line + 6, 1, address[1], highlight_format)
            except:
                worksheet.write(domestic_shipment_line + 6, 1, '', highlight_format)
            worksheet.write(domestic_shipment_line + 6, 2, '', highlight_format)
            worksheet.write(domestic_shipment_line + 7, 1, read_field(shipment, 'fields', 'FacilityCountry'), highlight_format)
            worksheet.write(domestic_shipment_line + 7, 2, '', highlight_format)
            
            # fill in line items
            # title
            for i in range(12):
                worksheet.write(domestic_shipment_line + 8, i, '箱子规格 / case dimensions' if i == 7 else '', table_header_without_border_format)
            
            # headers
            worksheet.set_row(domestic_shipment_line + 9, 50)
            worksheet.write(domestic_shipment_line + 9, 0, 'SKU', table_header_format)
            worksheet.write(domestic_shipment_line + 9, 1, 'FNSKU 条形码编号', table_header_format)
            worksheet.write(domestic_shipment_line + 9, 2, 'Units per Case /外箱包装', table_header_format)
            worksheet.write(domestic_shipment_line + 9, 3, 'Number of Cases/箱数量', table_header_format)
            worksheet.write(domestic_shipment_line + 9, 4, 'Units Shipped 订货数量（套）', table_header_format)
            worksheet.write(domestic_shipment_line + 9, 5, 'Total KG 总公斤', table_header_format)
            worksheet.write(domestic_shipment_line + 9, 6, 'Total CBM 总立方米', table_header_format)
            worksheet.write(domestic_shipment_line + 9, 7, '长 length', table_header_format)
            worksheet.write(domestic_shipment_line + 9, 8, '宽 width', table_header_format)
            worksheet.write(domestic_shipment_line + 9, 9, '高 height', table_header_format)
            worksheet.write(domestic_shipment_line + 9, 10, '总CBM', table_header_format)
            worksheet.write(domestic_shipment_line + 9, 11, 'Weight Per Case 外箱重量', table_header_format)
            worksheet.write(domestic_shipment_line + 9, 12, 'Box Mark分箱号：', box_header_format)

            # line item values
            for index, line_item in enumerate(domestic_shipment['line_items']):
                sku = read_field(line_item, 'sku')
                packaging_profile = read_field(line_item, 'packaging_profile')

                worksheet.write(domestic_shipment_line_start + index, 0, read_field(sku, 'fields', 'SKU'), rect_format)
                worksheet.write(domestic_shipment_line_start + index, 1, read_field(sku ,'fields', 'FNSKU'), rect_format)
                worksheet.write(domestic_shipment_line_start + index, 2, read_field(packaging_profile, 'fields', 'UnitsPerCarton'), rect_integer_format)
                worksheet.write(domestic_shipment_line_start + index, 3, read_field(line_item, 'fields', 'CaseQty'), rect_integer_format)
                worksheet.write(domestic_shipment_line_start + index, 4, read_field(line_item, 'fields', 'ShipQuantity'), rect_integer_format)
                worksheet.write_formula(domestic_shipment_line_start + index, 5, '=L{}*D{}'.format(domestic_shipment_line_start + index + 1, domestic_shipment_line_start + index + 1), rect_number_format)
                worksheet.write_formula(domestic_shipment_line_start + index, 6, '=K{}*D{}'.format(domestic_shipment_line_start + index + 1, domestic_shipment_line_start + index + 1), rect_number_format)
                worksheet.write(domestic_shipment_line_start + index, 7, read_field(packaging_profile,'fields', 'CartonLengthCM'), rect_integer_format)
                worksheet.write(domestic_shipment_line_start + index, 8, read_field(packaging_profile,'fields', 'CartonWidthCM'), rect_integer_format)
                worksheet.write(domestic_shipment_line_start + index, 9, read_field(packaging_profile,'fields', 'CartonHeightCM'), rect_integer_format)
                worksheet.write_formula(domestic_shipment_line_start + index, 10, '=H{}*I{}*J{}/1000000'.format(domestic_shipment_line_start + index + 1, domestic_shipment_line_start + index + 1, domestic_shipment_line_start + index + 1), rect_number_format)
                worksheet.write(domestic_shipment_line_start + index, 11, read_field(packaging_profile, 'fields', 'CartonWeightKG'), rect_number_format)
                worksheet.write(domestic_shipment_line_start + index, 12, read_field(line_item, 'fields', 'BoxMark'), rect_box_format)

            # draw bottom thick border
            for i in range(12):
                worksheet.write(domestic_shipment_line_end + 1, i, '', border_thick_top_format)

            domestic_shipment['total_line'] = domestic_shipment_line + 3
            domestic_shipment_line = domestic_shipment_line_end + 6
            domestic_shipment['domestic_shipment_line_start'] = domestic_shipment_line_start
            domestic_shipment['domestic_shipment_line_end'] = domestic_shipment_line_end


        # fill in sku information
        for index, sku in enumerate(skus):
            worksheet.write(3 + index, 0, sku)

            worksheet.write_formula(3 + index, 1, 
                '={}'.format(
                    '+'.join(
                        [
                            'SUMIF($A${}:$A${},$A${},$E${}:$E${})'.format(
                                domestic_shipment['domestic_shipment_line_start'] + 1,
                                domestic_shipment['domestic_shipment_line_end'] + 1,
                                4 + index,
                                domestic_shipment['domestic_shipment_line_start'] + 1,
                                domestic_shipment['domestic_shipment_line_end'] + 1
                            ) for domestic_shipment in domestic_shipments
                        ]
                    )
                ), integer_format
            )
            worksheet.write_formula(3 + index, 2, 
                '={}'.format(
                    '+'.join(
                        [
                            'SUMIF($A${}:$A${},$A${},$D${}:$D${})'.format(
                                domestic_shipment['domestic_shipment_line_start'] + 1,
                                domestic_shipment['domestic_shipment_line_end'] + 1,
                                4 + index,
                                domestic_shipment['domestic_shipment_line_start'] + 1,
                                domestic_shipment['domestic_shipment_line_end'] + 1
                            ) for domestic_shipment in domestic_shipments
                        ]
                    )
                ), integer_format
            )

        worksheet.write(3 + len(skus), 0, '', border_top_format)
        worksheet.write_formula(3 + len(skus), 1, '=SUM(B4:B{})'.format(3 + len(skus)), workbook.add_format({
            'top': 1,
            'num_format': '0'
        }))
        worksheet.write_formula(3 + len(skus), 2, '=SUM(C4:C{})'.format(3 + len(skus)), workbook.add_format({
            'top': 1,
            'num_format': '0'
        }))

        # fill in total information
        
        # Total KG
        worksheet.write(3, 5, 'Total KG', workbook.add_format({
            'fg_color': 'yellow',
            'top': 2,
            'left': 2
        }))
        worksheet.write_formula(3, 6, '={}'.format('+'.join([
            '$I${}'.format(
                domestic_shipment['total_line'] + 1
            ) for domestic_shipment in domestic_shipments
        ])), workbook.add_format({
            'fg_color': 'yellow',
            'num_format': '0.00',
            'right': 2,
            'top': 2
        }))

        # Total CBM
        worksheet.write(4, 5, 'Total CBM', workbook.add_format({
            'fg_color': 'yellow',
            'left': 2
        }))
        worksheet.write_formula(4, 6, '={}'.format('+'.join([
            '$I${}'.format(
                domestic_shipment['total_line'] + 2
            ) for domestic_shipment in domestic_shipments
        ])), workbook.add_format({
            'fg_color': 'yellow',
            'num_format': '0.00',
            'right': 2
        }))

        # Total Carto
        worksheet.write(5, 5, 'Total Cartons', workbook.add_format({
            'fg_color': 'yellow',
            'bottom': 2,
            'left': 2
        }))
        worksheet.write_formula(5, 6, '={}'.format('+'.join([
            '$I${}'.format(
                domestic_shipment['total_line'] + 3
            ) for domestic_shipment in domestic_shipments
        ])), workbook.add_format({
            'fg_color': 'yellow',
            'right': 2,
            'bottom': 2,
            'num_format': '0',
        }))

        workbook.close()
        
        print('##### Generating packaging list finished #####')
        return output.getvalue()
    except Exception as e:
        print('Error generating packaging list: ' + str(e))
        raise ValueError('Error generating packaging list: ' + str(e))


def create(event, context):
    print("Request Body: ")
    print(event["body"])
    
    try:
        body = json.loads(event["body"])
    except Exception as e:
        print(e)
        return {
            "statusCode": 500,
            "headers": {
                "Access-Control-Allow-Origin": "*"
            },
            "body": json.dumps({
                "error": "Error occured",
                "message": str(e)
            })
        }

    try:
        import boto3
        
        s3_client = boto3.client('s3')
        object_name = '{}.xlsx'.format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        
        packaging_list = generate_excel_file(
            get_domestic_shipments_from_airtable(os.getenv('AIRTABLE_APP_ID'), os.getenv('AIRTABLE_SECRET_KEY'), body['recordId'])
        )

        print('##### Putting generated packaging list to S3 started. Object name: ', object_name, ' #####')
        s3_client.put_object(
            Body = packaging_list,
            Bucket = os.getenv('BUCKET_NAME'),
            Key = object_name,
            ACL ='public-read'
        )
        print('##### Putting generated packaging list to S3 finished #####')
        
        upload_packaging_list_to_airtable(
            os.getenv('AIRTABLE_APP_ID'),
            os.getenv('AIRTABLE_SECRET_KEY'),
            body['recordId'],
            "https://{}.s3.amazonaws.com/{}".format(os.getenv('BUCKET_NAME'), object_name)
        )

        return {
            "statusCode": 200,
            "headers": {
                "Access-Control-Allow-Origin": "*"
            },
            "body": json.dumps({
                "message": "A new packaging list is generated and attached to the Airtable",
                "download": "https://{}.s3.amazonaws.com/{}".format(os.getenv('BUCKET_NAME'), object_name)
            })
        }
    except Exception as e:
        print(e)
        return {
            "statusCode": 500,
            "headers": {
                "Access-Control-Allow-Origin": "*"
            },
            "body": json.dumps({
                "error": "Error occured",
                "message": str(e)
            })
        }
