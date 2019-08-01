import os
import cx_Oracle
import traceback
import xlsxwriter
from flask import (
    Flask, Blueprint,
    send_from_directory
)
from datetime import datetime


app = Flask(__name__)
stores = Blueprint('stores', __name__, url_prefix='/maintain/storesExport', static_folder='static')

username = os.getenv('ORCL_USERNAME') or 'username'
password = os.getenv('ORCL_PASSWORD') or 'password'
dbUrl = os.getenv('ORCL_DBURL') or '127.0.0.1:1521/orcl'


def executeSql(sql, **kw):
    con = cx_Oracle.connect(username, password, dbUrl)
    cursor = con.cursor()
    result = None
    try:
        cursor.prepare(sql)
        cursor.execute(None, kw)
        result = cursor.fetchall()
        con.commit()
    except Exception:
        traceback.print_exc()
        con.rollback()
    finally:
        cursor.close()
        con.close()
    return result


@stores.route('/export/<itemNo>', methods=['GET'])
def reduceRecords(itemNo):
    now = datetime.now()
    xlsxDir = "export"
    fileName = "{}_{}_stores.xlsx".format(itemNo, now.strftime('%Y%m%d%H%M%S'))
    print("fileName is: {}".format(fileName))
    wb = xlsxwriter.Workbook(os.path.join(xlsxDir, fileName))
    sql = '''
    select t.item_no, t1.invt_no, '总署进境清单', t.legal_o_qty, time_to_char(t.input_date)
    from store_bill_goods_list t
    left outer join ceb2_invt_head t1 on t1.head_guid = t.seq_no
    where t.new_bill_type = '1'
    and t.legal_o_qty > 0
    and t.item_no = :itemNo
    union all
    select t.item_no, t.seq_no, (case substr(t.seq_no, 0, 1)
    when 'B' then '报关申请单'
    when 'D' then '调拨申请单'
    when 'K' then '库存调整单'
    when 'E' then '出区抽检'
    when 'T' then '一线退货出区'
    end), t.legal_o_qty, time_to_char(t.input_date) from store_bill_goods_list t
    where t.legal_o_qty > 0
    and (t.new_bill_type is null or t.new_bill_type = '3')
    and t.item_no = :itemNo
    '''
    result = executeSql(sql, itemNo=itemNo)
    sht1 = wb.add_worksheet('出区记录')
    sht1.write_string(0, 0, '商品备案编号')
    sht1.write_string(0, 1, '单证编号')
    sht1.write_string(0, 2, '单证类型')
    sht1.write_string(0, 3, '出区数量')
    sht1.write_string(0, 4, '出区时间')

    row = 1
    for invt in result:
        sht1.write(row, 0, invt[0])
        sht1.write(row, 1, invt[1])
        sht1.write(row, 2, invt[2])
        sht1.write(row, 3, invt[3])
        sht1.write(row, 4, invt[4])
        row += 1

    sql = '''
    select t.item_no, nvl(t.bill_no, t1.pre_no), (case substr(t.bill_no, 0, 1)
    when 'B' then '报关申请单'
    when 'D' then '调拨申请单'
    when 'K' then '库存调整单'
    when 'T' then '总署退货申请单'
    else '总署退货申请单'
    end), t.legal_i_qty, time_to_char(t.input_date) from store_bill_goods_list t
    left outer join ceb2_invt_refund_head t1 on t1.head_guid = t.seq_no
    where t.legal_i_qty > 0
    and t.item_no = :itemNo
    '''
    result = executeSql(sql, itemNo=itemNo)

    sht2 = wb.add_worksheet('入区记录')
    sht2.write_string(0, 0, '商品备案编号')
    sht2.write_string(0, 1, '单证编号')
    sht2.write_string(0, 2, '单证类型')
    sht2.write_string(0, 3, '出区数量')
    sht2.write_string(0, 4, '出区时间')

    row = 1
    for invt in result:
        sht2.write(row, 0, invt[0])
        sht2.write(row, 1, invt[1])
        sht2.write(row, 2, invt[2])
        sht2.write(row, 3, invt[3])
        sht2.write(row, 4, invt[4])
        row += 1

    wb.close()
    return send_from_directory(xlsxDir, fileName, as_attachment=True)


app.register_blueprint(stores)
