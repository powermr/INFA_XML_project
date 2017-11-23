# -*- conding:utf-8 -*-

import sys
import os
import glob
import re

import xml.dom.minidom as Dom
from xlrd import open_workbook
from xlwt import Workbook

import get_field_info


def read_xlsx(excel_file):
    '''
    读取模型-贴源对应关系excel文件，
    返回字典
    字典key为模型表名
    字典value为[m_field_name,m_table_type,s_table_name,s_field_name,s_table_type,s_dbname,m_dbname]
    '''
    dic = {}

    with open_workbook(excel_file) as workbook:
        worksheet = workbook.sheet_by_name('Sheet1')
        for row_index in range(2, worksheet.nrows):
            row = worksheet.row_values(row_index)
            dic.setdefault(row[0], []).append(row[1:])
    return dic


def addext_targetfield_node(doc_obj, parentNode, fn):
    # 属性列表
    attribut_list = [ \
        {'datatype': 'timestamp', 'fieldname': 'extend_field_time_stamp', 'precision': '29', 'scale': '9'}, \
        {'datatype': 'varchar', 'fieldname': 'extend_field_src_system', 'precision': '96', 'scale': '0'}, \
        {'datatype': 'integer', 'fieldname': 'extend_field_valid_flag', 'precision': '10', 'scale': '0'}, \
        {'datatype': 'integer', 'fieldname': 'extend_field_update_flag', 'precision': '10', 'scale': '0'}, \
        {'datatype': 'timestamp', 'fieldname': 'extend_field_update_time', 'precision': '29', 'scale': '9'}, \
        ]

    for i in range(1, 6):
        target_node = doc_obj.createElement('TARGETFIELD')
        datatype = attribut_list[i - 1]['datatype']
        fieldnumber = str(fn + i)
        fieldname = attribut_list[i - 1]['fieldname']
        precision = attribut_list[i - 1]['precision']
        scale = attribut_list[i - 1]['scale']

        target_node.setAttribute('BUSINESSNAME', '')
        target_node.setAttribute('DATATYPE', datatype)
        target_node.setAttribute('DESCRIPTION', '')
        target_node.setAttribute('FIELDNUMBER', fieldnumber)
        target_node.setAttribute('KEYTYPE', 'NOT A KEY')
        target_node.setAttribute('NAME', fieldname)
        target_node.setAttribute('NULLABLE', 'NULL')
        target_node.setAttribute('PICTURETEXT', '')
        target_node.setAttribute('PRECISION', precision)
        target_node.setAttribute('SCALE', scale)

        parentNode.appendChild(target_node)


def addext_targetfield_node(doc_obj, parentNode, fn):
    attribut_list = [ \
        {'datatype': 'varchar', 'fieldname': 'informatica_row_id', 'precision': '20', 'scale': '0'}, \
        {'datatype': 'varchar', 'fieldname': 'informatica_flag', 'precision': '10', 'scale': '0'}, \
        {'datatype': 'bigint', 'fieldname': 'EXT_OGG_SEQ', 'precision': '19', 'scale': '0'}, \
        {'datatype': 'varchar', 'fieldname': 'lineseparator', 'precision': '10', 'scale': '0'} \
        ]
    for i in range(1, 5):
        targetfield_node = doc_obj.createElement('TARGETFIELD')
        datatype = attribut_list[i - 1]['datatype']
        fieldnumber = str(fn + i)
        fieldname = attribut_list[i - 1]['fieldname']
        precision = attribut_list[i - 1]['precision']
        scale = attribut_list[i - 1]['scale']

        targetfield_node.setAttribute('BUSINESSNAME', '')
        targetfield_node.setAttribute('DATATYPE', datatype)
        targetfield_node.setAttribute('DESCRIPTION', '')
        targetfield_node.setAttribute('FIELDNUMBER', fieldnumber)
        targetfield_node.setAttribute('KEYTYPE', 'NOT A KEY')
        targetfield_node.setAttribute('NAME', fieldname)
        targetfield_node.setAttribute('NULLABLE', 'NULL')
        targetfield_node.setAttribute('PICTURETEXT', '')
        targetfield_node.setAttribute('PRECISION', precision)
        targetfield_node.setAttribute('SCALE', scale)

        parentNode.appendChild(targetfield_node)


def addext_sourcefield_node(doc_obj, parentNode, fn):
    # 属性列表
    attribut_list = [ \
        {'datatype': 'varchar', 'fieldname': 'informatica_row_id', 'precision': '20', 'scale': '0'}, \
        {'datatype': 'varchar', 'fieldname': 'informatica_flag', 'precision': '10', 'scale': '0'}, \
        {'datatype': 'timestamp', 'fieldname': 'informatica_date_time', 'precision': '29', 'scale': '9'}, \
        {'datatype': 'bigint', 'fieldname': 'EXT_OGG_SEQ', 'precision': '19', 'scale': '0'}, \
        {'datatype': 'varchar', 'fieldname': 'informatica_flag1', 'precision': '10', 'scale': '0'} \
        ]

    for i in range(1, 6):
        sourcefield_node = doc_obj.createElement('SOURCEFIELD')
        datatype = attribut_list[i - 1]['datatype']
        fieldnumber = str(fn + i)
        fieldname = attribut_list[i - 1]['fieldname']
        precision = attribut_list[i - 1]['precision']
        scale = attribut_list[i - 1]['scale']

        sourcefield_node.setAttribute('BUSINESSNAME', '')
        sourcefield_node.setAttribute('DATATYPE', datatype)
        sourcefield_node.setAttribute('DESCRIPTION', '')
        sourcefield_node.setAttribute('FIELDNUMBER', fieldnumber)
        sourcefield_node.setAttribute('FIELDPROPERTY', '0')
        sourcefield_node.setAttribute('FIELDTYPE', 'ELEMITEM')
        sourcefield_node.setAttribute('HIDDEN', 'NO')
        sourcefield_node.setAttribute('KEYTYPE', 'NOT A KEY')
        sourcefield_node.setAttribute('LENGTH', '0')
        sourcefield_node.setAttribute('LEVEL', '0')
        sourcefield_node.setAttribute('NAME', fieldname)
        sourcefield_node.setAttribute('NULLABLE', 'NULL')
        sourcefield_node.setAttribute('OCCURS', '0')
        sourcefield_node.setAttribute('OFFSET', '0')
        sourcefield_node.setAttribute('PHYSICALLENGTH', '50')
        sourcefield_node.setAttribute('PHYSICALOFFSET', '0')
        sourcefield_node.setAttribute('PICTURETEXT', '')
        sourcefield_node.setAttribute('PRECISION', precision)
        sourcefield_node.setAttribute('SCALE', scale)
        sourcefield_node.setAttribute('USAGE_FLAGS', '')

        parentNode.appendChild(sourcefield_node)


def addext_sq_transformfield_node(doc_obj, parentNode):
    attribut_list = [ \
        {'datatype': 'string', 'fieldname': 'informatica_row_id', 'precision': '20', 'scale': '0'}, \
        {'datatype': 'string', 'fieldname': 'informatica_flag', 'precision': '10', 'scale': '0'}, \
        {'datatype': 'date/time', 'fieldname': 'informatica_date_time', 'precision': '29', 'scale': '9'}, \
        {'datatype': 'bigint', 'fieldname': 'EXT_OGG_SEQ', 'precision': '19', 'scale': '0'}, \
        {'datatype': 'string', 'fieldname': 'informatica_flag1', 'precision': '10', 'scale': '0'} \
        ]
    for i in range(1, 6):
        tranfield_node = doc_obj.createElement('TRANSFORMFIELD')

        datatype = attribut_list[i - 1]['datatype']
        fieldname = attribut_list[i - 1]['fieldname']
        precision = attribut_list[i - 1]['precision']
        scale = attribut_list[i - 1]['scale']

        tranfield_node.setAttribute('DATATYPE', datatype)
        tranfield_node.setAttribute('DEFAULTVALUE', '')
        tranfield_node.setAttribute('DESCRIPTION', '')
        tranfield_node.setAttribute('NAME', fieldname)
        tranfield_node.setAttribute('PICTURETEXT', '')
        tranfield_node.setAttribute('PORTTYPE', 'INPUT/OUTPUT')
        tranfield_node.setAttribute('PRECISION', precision)
        tranfield_node.setAttribute('SCALE', scale)

        parentNode.appendChild(tranfield_node)


def addext_ex_transformfield_node(doc_obj, parentNode, sysname):
    attr_list = [ \
        {'datatype': 'string', 'expression': 'informatica_row_id', 'fieldname': 'informatica_row_id',
         'porttype': 'INPUT/OUTPUT', 'precision': '20', 'scale': '0', 'defaultvalue': ''}, \
        {'datatype': 'string', 'expression': 'informatica_flag', 'fieldname': 'informatica_flag',
         'porttype': 'INPUT/OUTPUT', 'precision': '10', 'scale': '0', 'defaultvalue': ''}, \
        {'datatype': 'date/time', 'expression': 'informatica_date_time', 'fieldname': 'informatica_date_time',
         'porttype': 'INPUT/OUTPUT', 'precision': '29', 'scale': '9', 'defaultvalue': ''}, \
        {'datatype': 'bigint', 'expression': 'EXT_OGG_SEQ', 'fieldname': 'EXT_OGG_SEQ', 'porttype': 'INPUT/OUTPUT',
         'precision': '19', 'scale': '0', 'defaultvalue': ''}, \
        {'datatype': 'string', 'expression': '\'' + sysname + '\'', 'fieldname': 'extend_field_src_system',
         'porttype': 'OUTPUT', 'precision': '10', 'scale': '0', 'defaultvalue': 'ERROR(\'transformation error\')'}, \
        {'datatype': 'integer', 'expression': 'extend_field_valid_flag', 'fieldname': 'extend_field_valid_flag',
         'porttype': 'INPUT/OUTPUT', 'precision': '10', 'scale': '0', 'defaultvalue': ''}, \
        {'datatype': 'integer', 'expression': 'decode(informatica_flag,\'I\',0,\'D\',2,\'UI\',1,\'UD\',3)',
         'fieldname': 'extend_field_update_flag', 'porttype': 'OUTPUT', 'precision': '10', 'scale': '0',
         'defaultvalue': 'ERROR(\'transformation error\')'}, \
        {'datatype': 'string', 'expression': 'SETVARIABLE($$date,TO_CHAR(SESSSTARTTIME,\'YYYY/MM/DD\'))',
         'fieldname': '$$date', 'porttype': 'LOCAL VARIABLE', 'precision': '10', 'scale': '0', 'defaultvalue': ''}, \
        {'datatype': 'string', 'expression': '\'(a#a)\'', 'fieldname': 'lineseparator', 'porttype': 'OUTPUT',
         'precision': '10', 'scale': '0', 'defaultvalue': ''} \
        ]

    for i in range(1, 10):
        tranfield_node = doc_obj.createElement('TRANSFORMFIELD')

        datatype = attr_list[i - 1]['datatype']
        defaultvalue = attr_list[i - 1]['defaultvalue']
        expression = attr_list[i - 1]['expression']
        fieldname = attr_list[i - 1]['fieldname']
        porttype = attr_list[i - 1]['porttype']
        precision = attr_list[i - 1]['precision']
        scale = attr_list[i - 1]['scale']

        tranfield_node.setAttribute('DATATYPE', datatype)
        tranfield_node.setAttribute('DEFAULTVALUE', defaultvalue)
        tranfield_node.setAttribute('DESCRIPTION', '')
        tranfield_node.setAttribute('EXPRESSION', expression)
        tranfield_node.setAttribute('EXPRESSIONTYPE', 'GENERAL')
        tranfield_node.setAttribute('NAME', fieldname)
        tranfield_node.setAttribute('PICTURETEXT', '')
        tranfield_node.setAttribute('PORTTYPE', porttype)
        tranfield_node.setAttribute('PRECISION', precision)
        tranfield_node.setAttribute('SCALE', scale)

        parentNode.appendChild(tranfield_node)
    tableattribute_node = doc_obj.createElement('TABLEATTRIBUTE')
    tableattribute_node.setAttribute('NAME', 'Tracing Level')
    tableattribute_node.setAttribute('VALUE', 'Normal')
    parentNode.appendChild(tableattribute_node)


def add_instance_node(doc_obj, parentNode, m_tbname, s_tb_name, s_db_name):
    instance_node_1 = doc_obj.createElement('INSTANCE')
    instance_node_1.setAttribute('DESCRIPTION', '')
    instance_node_1.setAttribute('NAME', m_tbname)
    instance_node_1.setAttribute('TRANSFORMATION_NAME', m_tbname)
    instance_node_1.setAttribute('TRANSFORMATION_TYPE', 'Target Definition')
    instance_node_1.setAttribute('TYPE', 'TARGET')
    parentNode.appendChild(instance_node_1)

    instance_node_2 = doc_obj.createElement('INSTANCE')
    instance_node_2.setAttribute('DBDNAME', 'mysql' + s_db_name)
    instance_node_2.setAttribute('DESCRIPTION', '')
    instance_node_2.setAttribute('NAME', s_tb_name)
    instance_node_2.setAttribute('TRANSFORMATION_NAME', s_tb_name)
    instance_node_2.setAttribute('TRANSFORMATION_TYPE', 'Source Definition')
    instance_node_2.setAttribute('TYPE', 'SOURCE')
    parentNode.appendChild(instance_node_2)

    instance_node_3 = doc_obj.createElement('INSTANCE')
    instance_node_3.setAttribute('NAME', 'SQ_' + s_tb_name)
    instance_node_3.setAttribute('REUSABLE', 'NO')
    instance_node_3.setAttribute('TRANSFORMATION_NAME', 'SQ_' + s_tb_name)
    instance_node_3.setAttribute('TRANSFORMATION_TYPE', 'Source Qualifier')
    instance_node_3.setAttribute('TYPE', 'TRANSFORMATION')
    associated_source_instance_node = doc_obj.createElement('ASSOCIATED_SOURCE_INSTANCE')
    associated_source_instance_node.setAttribute('NAME', s_tb_name)
    instance_node_3.appendChild(associated_source_instance_node)
    parentNode.appendChild(instance_node_3)

    instance_node_4 = doc_obj.createElement('INSTANCE')
    instance_node_4.setAttribute('DESCRIPTION', '')
    instance_node_4.setAttribute('NAME', 'EXPTRANS')
    instance_node_4.setAttribute('REUSABLE', 'NO')
    instance_node_4.setAttribute('TRANSFORMATION_NAME', 'EXPTRANS')
    instance_node_4.setAttribute('TRANSFORMATION_TYPE', 'Expression')
    instance_node_4.setAttribute('TYPE', 'TRANSFORMATION')
    parentNode.appendChild(instance_node_4)


def create_connector_exptrans(doc_obj, fromfield, tofield, toinstance, lst):
    '''
    表达式到目标（模型）的连线
    fromfield：表达式字段，此处取source字段
    tofield：模型字段  target_field
    toinstance  目标实例，此处取模型表名target
    lst存放生成的connector节点对象
    '''
    connector_node = doc_obj.createElement('CONNECTOR')
    connector_node.setAttribute('FROMFIELD', fromfield)
    connector_node.setAttribute('FROMINSTANCE', 'EXPTRANS')
    connector_node.setAttribute('FROMINSTANCETYPE', 'Expression')
    connector_node.setAttribute('TOFIELD', tofield)
    connector_node.setAttribute('TOINSTANCE', toinstance)
    connector_node.setAttribute('TOINSTANCETYPE', 'Target Definition')
    lst.append(connector_node)


def create_connector_source(doc_obj, fromfield, frominstance, tofield, toinstance, lst):
    '''
    源到SQ组件的连线
    fromfield：源source字段
    frominstance:来源实例
    tofield：模型字段  target_field
    toinstance  目标实例，此处取模型表名target
    lst存放生成的connector节点对象
    '''
    connector_node = doc_obj.createElement('CONNECTOR')
    connector_node.setAttribute('FROMFIELD', fromfield)
    connector_node.setAttribute('FROMINSTANCE', frominstance)
    connector_node.setAttribute('FROMINSTANCETYPE', 'Source Definition')
    connector_node.setAttribute('TOFIELD', tofield)
    connector_node.setAttribute('TOINSTANCE', toinstance)
    connector_node.setAttribute('TOINSTANCETYPE', 'Source Qualifier')
    lst.append(connector_node)


def create_connector_sq(doc_obj, fromfield, frominstance, tofield, lst):
    '''
    sq到表达式的连线
    fromfield：表达式字段，此处取source字段
    tofield：模型字段  target_field
    toinstance  目标实例，此处取模型表名target
    lst存放生成的connector节点对象
    '''
    connector_node = doc_obj.createElement('CONNECTOR')
    connector_node.setAttribute('FROMFIELD', fromfield)
    connector_node.setAttribute('FROMINSTANCE', frominstance)
    connector_node.setAttribute('FROMINSTANCETYPE', 'Source Qualifier')
    connector_node.setAttribute('TOFIELD', tofield)
    connector_node.setAttribute('TOINSTANCE', 'EXPTRANS')
    connector_node.setAttribute('TOINSTANCETYPE', 'Expression')
    lst.append(connector_node)


def create_tag(document, s_tbl_name, s_db, info_dic, m_tbl_name, tbl_count, m_db, config):
    '''
        按模型表名为维度调用此模块
        document:document对象
        s_tbl_name:源表名
        s_db：源库名
        m_db：模型库名
        info_dic：字典类型，可根据m_tbl_name获取模型的字段名数据类型及其对应的source表名字段名数据类型
        m_tbl_name:模型表名
        tbl_count:表计数，用于生成mapping的name
        创建source标签，创建target标签，创建mapping标签
    
    '''
    system_name = str(info_dic[m_tbl_name][0][7])

    sq_name = 'SQ_' + s_tbl_name
    field_list = info_dic[m_tbl_name]
    mapping_name = 'm_' + str(tbl_count) + '_' + m_tbl_name

    folder_node = document.getElementsByTagName('FOLDER')[0]
    source_node = document.createElement('SOURCE')
    mapping_node = document.createElement('MAPPING')

    source_node.setAttribute('BUSINESSNAME', '')
    source_node.setAttribute('DATABASETYPE', 'ODBC')
    source_node.setAttribute('DBDNAME', 'mysql' + s_db)
    source_node.setAttribute('DESCRIPTION', '')
    source_node.setAttribute('NAME', s_tbl_name)
    source_node.setAttribute('OBJECTVERSION', '1')
    source_node.setAttribute('OWNERNAME', '')
    source_node.setAttribute('VERSIONNUMBER', '1')

    target_node = document.createElement('TARGET')
    target_node.setAttribute('BUSINESSNAME', '')
    target_node.setAttribute('CONSTRAINT', '')
    target_node.setAttribute('DATABASETYPE', 'ODBC')
    target_node.setAttribute('DESCRIPTION', '')
    target_node.setAttribute('NAME', m_tbl_name)
    target_node.setAttribute('OBJECTVERSION', '1')
    target_node.setAttribute('TABLEOPTIONS', '')
    target_node.setAttribute('VERSIONNUMBER', '1')

    mapping_node.setAttribute('DESCRIPTION', '')
    mapping_node.setAttribute('ISVALID', 'YES')
    mapping_node.setAttribute('NAME', mapping_name)
    mapping_node.setAttribute('OBJECTVERSION', '1')
    mapping_node.setAttribute('VERSIONNUMBER', '1')

    # 生成mapping_node的子标签sq组件的transformation
    transformation_sq_node = document.createElement('TRANSFORMATION')
    transformation_sq_node.setAttribute('DESCRIPTION', '')
    transformation_sq_node.setAttribute('NAME', sq_name)
    transformation_sq_node.setAttribute('OBJECTVERSION', '1')
    transformation_sq_node.setAttribute('REUSABLE', 'NO')
    transformation_sq_node.setAttribute('TYPE', 'Source Qualifier')
    transformation_sq_node.setAttribute('VERSIONNUMBER', '1')

    # 生成mapping_node的子标签expression组件的transformation
    transformation_ex_node = document.createElement('TRANSFORMATION')
    transformation_ex_node.setAttribute('DESCRIPTION', '')
    transformation_ex_node.setAttribute('NAME', 'EXPTRANS')
    transformation_ex_node.setAttribute('OBJECTVERSION', '1')
    transformation_ex_node.setAttribute('REUSABLE', 'NO')
    transformation_ex_node.setAttribute('TYPE', 'Expression')
    transformation_ex_node.setAttribute('VERSIONNUMBER', '1')

    fieldnumber = 0
    m_fieldnumber = 0
    sql_query_string = ''
    load_sql_query_string = ''
    conn_list = []

    for fields in field_list:

        source_field = fields[3]
        source_field_type = fields[4]

        target_field = fields[0]
        target_field_type = fields[1]

        if not target_field == '':
            target_field_info = get_field_info.source_field_type(
                target_field_type)
            target_field_datatype = target_field_info[0]
            target_field_precision = target_field_info[1]
            target_field_scale = target_field_info[2]

            m_fieldnumber = m_fieldnumber + 1
            targetfield_node = document.createElement('TARGETFIELD')
            targetfield_node.setAttribute('BUSINESSNAME', '')
            targetfield_node.setAttribute('DATATYPE', target_field_datatype)
            targetfield_node.setAttribute('DESCRIPTION', '')
            targetfield_node.setAttribute('FIELDNUMBER', str(m_fieldnumber))
            targetfield_node.setAttribute('KEYTYPE', 'NOT A KEY')
            targetfield_node.setAttribute('NAME', target_field)
            targetfield_node.setAttribute('NULLABLE', 'NULL')
            targetfield_node.setAttribute('PICTURETEXT', '')
            targetfield_node.setAttribute('PRECISION', target_field_precision)
            targetfield_node.setAttribute('SCALE', target_field_scale)

            target_node.appendChild(targetfield_node)

        if not source_field == '':
            # 添加拆分source_field_type的处理
            field_info = get_field_info.source_field_type(source_field_type)
            source_field_datatype = field_info[0]
            source_field_precision = field_info[1]
            source_field_scale = field_info[2]

            fieldnumber = fieldnumber + 1

            # 创建sourcefield标签
            sourcefield_node = document.createElement('SOURCEFIELD')
            sourcefield_node.setAttribute('BUSINESSNAME', '')
            sourcefield_node.setAttribute('DATATYPE', source_field_datatype)
            sourcefield_node.setAttribute('DESCRIPTION', '')
            sourcefield_node.setAttribute('FIELDNUMBER', str(fieldnumber))
            sourcefield_node.setAttribute('FIELDPROPERTY', '0')
            sourcefield_node.setAttribute('FIELDTYPE', 'ELEMITEM')
            sourcefield_node.setAttribute('HIDDEN', 'NO')
            sourcefield_node.setAttribute('KEYTYPE', 'NOT A KEY')
            sourcefield_node.setAttribute('LENGTH', '0')
            sourcefield_node.setAttribute('LEVEL', '0')
            sourcefield_node.setAttribute('NAME', source_field)
            sourcefield_node.setAttribute('NULLABLE', 'NULL')
            sourcefield_node.setAttribute('OCCURS', '0')
            sourcefield_node.setAttribute('OFFSET', '0')
            sourcefield_node.setAttribute('PHYSICALLENGTH', '50')
            sourcefield_node.setAttribute('PHYSICALOFFSET', '0')
            sourcefield_node.setAttribute('PICTURETEXT', '')
            sourcefield_node.setAttribute('PRECISION', source_field_precision)
            sourcefield_node.setAttribute('SCALE', source_field_scale)
            sourcefield_node.setAttribute('USAGE_FLAGS', '')

            source_node.appendChild(sourcefield_node)

            # 创建SQ组件的transformfield标签
            sq_transformfield_node = document.createElement('TRANSFORMFIELD')

            # 创建EX的transformfield
            ex_transformfield_node = document.createElement('TRANSFORMFIELD')

            '''====================================
            #这里的类型和精度需要确认：mapping里的字段类型应该对应目标还是源
            sq_datatype=""
            source_field_precision=""
            source_field_scale=""
            ===================================='''
            ex_datatype_info = get_field_info.s_sq_type_transform(field_info)
            ex_datatype = ex_datatype_info[0]
            ex_field_precision = ex_datatype_info[1]
            ex_field_scale = ex_datatype_info[2]

            sq_transformfield_node.setAttribute('DATATYPE', ex_datatype)
            ex_transformfield_node.setAttribute('DATATYPE', ex_datatype)

            sq_transformfield_node.setAttribute('DEFAULTVALUE', '')
            ex_transformfield_node.setAttribute('DEFAULTVALUE', '')

            sq_transformfield_node.setAttribute('DESCRIPTION', '')
            ex_transformfield_node.setAttribute('DESCRIPTION', '')

            sq_transformfield_node.setAttribute('NAME', source_field)
            ex_transformfield_node.setAttribute('EXPRESSION', source_field)
            ex_transformfield_node.setAttribute('EXPRESSIONTYPE', 'GENERAL')
            ex_transformfield_node.setAttribute('NAME', source_field)

            sq_transformfield_node.setAttribute('PICTURETEXT', '')
            ex_transformfield_node.setAttribute('PICTURETEXT', '')

            sq_transformfield_node.setAttribute('PORTTYPE', 'INPUT/OUTPUT')
            ex_transformfield_node.setAttribute('PORTTYPE', 'INPUT/OUTPUT')

            sq_transformfield_node.setAttribute('PRECISION', ex_field_precision)
            ex_transformfield_node.setAttribute('PRECISION', ex_field_precision)

            sq_transformfield_node.setAttribute('SCALE', ex_field_scale)
            ex_transformfield_node.setAttribute('SCALE', ex_field_scale)

            transformation_sq_node.appendChild(sq_transformfield_node)
            transformation_ex_node.appendChild(ex_transformfield_node)

        # 创建连线
        if not target_field == '' and not source_field == '':
            create_connector_exptrans(document, source_field, target_field, m_tbl_name, conn_list)
            create_connector_source(document, source_field, s_tbl_name, source_field, sq_name, conn_list)
            create_connector_sq(document, source_field, sq_name, source_field, conn_list)

            sql_query_string = sql_query_string + 'c.' + source_field + ','
            load_sql_query_string = load_sql_query_string + source_field + ','

    # 补充create_connector_exptrans 中剩余的固定连线
    # sq 到表达式的
    sq_to_exp_arr = [ \
        {"FROMFIELD": "informatica_date_time", "TOFIELD": "informatica_date_time"}, \
        {"FROMFIELD": "informatica_row_id", "TOFIELD": "informatica_row_id"}, \
        {"FROMFIELD": "informatica_flag", "TOFIELD": "informatica_flag"}, \
        {"FROMFIELD": "EXT_OGG_SEQ", "TOFIELD": "EXT_OGG_SEQ"}, \
        {"FROMFIELD": "informatica_flag1", "TOFIELD": "extend_field_valid_flag"} \
        ]
    for sq_to_exp in sq_to_exp_arr:
        create_connector_sq(document, sq_to_exp["FROMFIELD"], sq_name, sq_to_exp["TOFIELD"], conn_list)
    # 源到sq
    s_to_sq_arr = [ \
        {"FROMFIELD": "informatica_date_time", "TOFIELD": "informatica_date_time"}, \
        {"FROMFIELD": "informatica_row_id", "TOFIELD": "informatica_row_id"}, \
        {"FROMFIELD": "informatica_flag", "TOFIELD": "informatica_flag"}, \
        {"FROMFIELD": "EXT_OGG_SEQ", "TOFIELD": "EXT_OGG_SEQ"}, \
        {"FROMFIELD": "informatica_flag1", "TOFIELD": "informatica_flag1"} \
        ]
    for s_to_sq in s_to_sq_arr:
        create_connector_source(document, s_to_sq["FROMFIELD"], s_tbl_name, s_to_sq["TOFIELD"], sq_name, conn_list)

    # 表达式到目标（模型）
    exp_to_m_arr = [ \
        {"FROMFIELD": "informatica_date_time", "TOFIELD": "extend_field_update_time"}, \
        {"FROMFIELD": "informatica_row_id", "TOFIELD": "informatica_row_id"}, \
        {"FROMFIELD": "informatica_flag", "TOFIELD": "informatica_flag"}, \
        {"FROMFIELD": "EXT_OGG_SEQ", "TOFIELD": "ext_ogg_seq"}, \
        {"FROMFIELD": "extend_field_src_system", "TOFIELD": "extend_field_src_system"}, \
        {"FROMFIELD": "extend_field_valid_flag", "TOFIELD": "extend_field_valid_flag"}, \
        {"FROMFIELD": "extend_field_update_flag", "TOFIELD": "extend_field_update_flag"}, \
        {"FROMFIELD": "lineseparator", "TOFIELD": "lineseparator"} \
        ]
    for exp_to_m in exp_to_m_arr:
        create_connector_exptrans(document, exp_to_m["FROMFIELD"], exp_to_m["TOFIELD"], m_tbl_name, conn_list)

    # 添加excel中没有的扩展字段
    addext_targetfield_node(document, target_node, m_fieldnumber)
    addext_sourcefield_node(document, source_node, fieldnumber)

    # 添加sq组件和ex组件的扩展信息
    addext_sq_transformfield_node(document, transformation_sq_node)
    addext_ex_transformfield_node(document, transformation_ex_node, system_name)

    # 在transformation_sq_node节点上追加tableattribute节点
    sql_query = "SELECT " + \
                sql_query_string + \
                "c.informatica_row_id,c.informatica_flag,c.informatica_date_time,c.EXT_OGG_SEQ,CASE WHEN (" + \
                "c.informatica_flag = 'I' AND c.EXT_OGG_SEQ = C.maxEXT_OGG_SEQ) THEN	1 WHEN (" + \
                "c.informatica_flag = 'UI' AND c.EXT_OGG_SEQ = C.maxEXT_OGG_SEQ) THEN	1 ELSE	0 " + \
                "END AS valid_flag FROM(SELECT	a.*, b.maxEXT_OGG_SEQ FROM " + s_tbl_name + " a,(" + \
                "SELECT informatica_row_id,max(EXT_OGG_SEQ) maxEXT_OGG_SEQ FROM " + s_tbl_name + \
                " GROUP BY   informatica_row_id) b  WHERE  a.informatica_row_id = b.informatica_row_id ) c " + \
                "where c.informatica_date_time >= date_format('$$date','%y%m%d')" + \
                " and c.informatica_date_time < date_format('$$$SESSSTARTTIME','%y%m%d')"

    tbl_attri_list = [ \
        {'name': 'Sql Query', 'value': sql_query}, \
        {'name': 'User Defined Join', 'value': ''}, \
        {'name': 'Source Filter', 'value': ''}, \
        {'name': 'Number Of Sorted Ports', 'value': '0'}, \
        {'name': 'Tracing Level', 'value': 'Normal'}, \
        {'name': 'Select Distinct', 'value': 'NO'}, \
        {'name': 'Is Partitionable', 'value': 'NO'}, \
        {'name': 'Pre SQL', 'value': ''}, \
        {'name': 'Post SQL', 'value': ''}, \
        {'name': 'Output is deterministic', 'value': 'NO'}, \
        {'name': 'Output is repeatable', 'value': 'Never'} \
        ]
    for attr in tbl_attri_list:
        tableattribute = document.createElement('TABLEATTRIBUTE')
        tableattribute.setAttribute('NAME', attr['name'])
        tableattribute.setAttribute('VALUE', attr['value'])

        transformation_sq_node.appendChild(tableattribute)

    folder_node.appendChild(source_node)
    folder_node.appendChild(target_node)
    mapping_node.appendChild(transformation_sq_node)
    mapping_node.appendChild(transformation_ex_node)
    # 在mapping中添加INSTANCE标签
    add_instance_node(document, mapping_node, m_tbl_name, s_tbl_name, s_db)
    # 将conn_list中的connector标签追加到mapping节点
    for conn in conn_list:
        mapping_node.appendChild(conn)

    # 添加TARGETLOADORDER MAPPINGVARIABLE  ERPINFO 标签
    targetloadorder_node = document.createElement('TARGETLOADORDER')
    targetloadorder_node.setAttribute('ORDER', '1')
    targetloadorder_node.setAttribute('TARGETINSTANCE', m_tbl_name)
    mapping_node.appendChild(targetloadorder_node)

    mappingvariable_node = document.createElement('MAPPINGVARIABLE')
    mappingvariable_node.setAttribute('AGGFUNCTION', 'MAX')
    mappingvariable_node.setAttribute('DATATYPE', 'string')
    mappingvariable_node.setAttribute('DEFAULTVALUE', '2000/01/01')
    mappingvariable_node.setAttribute('ISEXPRESSIONVARIABLE', 'NO')
    mappingvariable_node.setAttribute('ISPARAM', 'NO')
    mappingvariable_node.setAttribute('NAME', '$$date')
    mappingvariable_node.setAttribute('PRECISION', '20')
    mappingvariable_node.setAttribute('SCALE', '0')
    mappingvariable_node.setAttribute('USERDEFINED', 'YES')
    mapping_node.appendChild(mappingvariable_node)

    erpinfo_node = document.createElement("ERPINFO")
    mapping_node.appendChild(erpinfo_node)

    folder_node.appendChild(mapping_node)

    # 添加config节点
    folder_node.appendChild(config)

    # 生成workflow_node节点
    workflow_name = 'wf_m_' + str(tbl_count) + '_' + m_tbl_name
    workflow_node = document.createElement('WORKFLOW')
    workflow_node.setAttribute('DESCRIPTION', '这些工作流是通过生成工作流向导创建的。')
    workflow_node.setAttribute('ISENABLED', 'YES')
    workflow_node.setAttribute('ISRUNNABLESERVICE', 'NO')
    workflow_node.setAttribute('ISSERVICE', 'NO')
    workflow_node.setAttribute('ISVALID', 'YES')
    workflow_node.setAttribute('NAME', workflow_name)
    workflow_node.setAttribute('REUSABLE_SCHEDULER', 'NO')
    workflow_node.setAttribute('SCHEDULERNAME', '计划程序')
    workflow_node.setAttribute('SERVERNAME', 'infa_rep_services_bak')
    workflow_node.setAttribute('SERVER_DOMAINNAME', 'Domain_infa')
    workflow_node.setAttribute('SUSPEND_ON_ERROR', 'NO')
    workflow_node.setAttribute('TASKS_MUST_RUN_ON_SERVER', 'NO')
    workflow_node.setAttribute('VERSIONNUMBER', '1')

    scheduler_node = document.createElement('SCHEDULER')
    scheduler_node.setAttribute('DESCRIPTION', '')
    scheduler_node.setAttribute('NAME', '计划程序')
    scheduler_node.setAttribute('REUSABLE', 'NO')
    scheduler_node.setAttribute('VERSIONNUMBER', '1')
    scheduleinfo_node = document.createElement('SCHEDULEINFO')
    scheduleinfo_node.setAttribute('SCHEDULETYPE', 'ONDEMAND')
    scheduler_node.appendChild(scheduleinfo_node)
    workflow_node.appendChild(scheduler_node)

    task_node = document.createElement('TASK')
    task_node.setAttribute('DESCRIPTION', '')
    task_node.setAttribute('NAME', '启动')
    task_node.setAttribute('REUSABLE', 'NO')
    task_node.setAttribute('TYPE', 'Start')
    task_node.setAttribute('VERSIONNUMBER', '1')
    workflow_node.appendChild(task_node)

    ############session_node##########################
    session_name = 's_m_' + str(tbl_count) + '_' + m_tbl_name

    session_node = document.createElement('SESSION')
    session_node.setAttribute('DESCRIPTION', '')
    session_node.setAttribute('ISVALID', 'YES')
    session_node.setAttribute('MAPPINGNAME', mapping_name)
    session_node.setAttribute('NAME', session_name)
    session_node.setAttribute('REUSABLE', 'NO')
    session_node.setAttribute('SORTORDER', 'Binary')
    session_node.setAttribute('VERSIONNUMBER', '1')

    sesstransformationinst_node = document.createElement('SESSTRANSFORMATIONINST')
    sesstransformationinst_node.setAttribute('ISREPARTITIONPOINT', 'YES')
    sesstransformationinst_node.setAttribute('PARTITIONTYPE', 'PASS THROUGH')
    sesstransformationinst_node.setAttribute('PIPELINE', '1')
    sesstransformationinst_node.setAttribute('SINSTANCENAME', m_tbl_name)
    sesstransformationinst_node.setAttribute('STAGE', '1')
    sesstransformationinst_node.setAttribute('TRANSFORMATIONNAME', m_tbl_name)
    sesstransformationinst_node.setAttribute('TRANSFORMATIONTYPE', 'Target Definition')
    flatfile_node = document.createElement('FLATFILE')
    flatfile_node.setAttribute('CODEPAGE', 'UTF-8')
    flatfile_node.setAttribute('CONSECDELIMITERSASONE', 'NO')
    flatfile_node.setAttribute('DELIMITED', 'YES')
    flatfile_node.setAttribute('DELIMITERS', '@#@')
    flatfile_node.setAttribute('ESCAPE_CHARACTER', '')
    flatfile_node.setAttribute('KEEPESCAPECHAR', 'NO')
    flatfile_node.setAttribute('LINESEQUENTIAL', 'NO')
    flatfile_node.setAttribute('MULTIDELIMITERSASAND', 'NO')
    flatfile_node.setAttribute('NULLCHARTYPE', 'ASCII')
    flatfile_node.setAttribute('NULL_CHARACTER', '*')
    flatfile_node.setAttribute('PADBYTES', '1')
    flatfile_node.setAttribute('QUOTE_CHARACTER', 'NONE')
    flatfile_node.setAttribute('REPEATABLE', 'NO')
    flatfile_node.setAttribute('ROWDELIMITER', '10')
    flatfile_node.setAttribute('SKIPROWS', '0')
    flatfile_node.setAttribute('STRIPTRAILINGBLANKS', 'NO')
    sesstransformationinst_node.appendChild(flatfile_node)
    session_node.appendChild(sesstransformationinst_node)

    sesstransformationinst_node = document.createElement('SESSTRANSFORMATIONINST')
    sesstransformationinst_node.setAttribute('ISREPARTITIONPOINT', 'NO')
    sesstransformationinst_node.setAttribute('PIPELINE', '0')
    sesstransformationinst_node.setAttribute('SINSTANCENAME', s_tbl_name)
    sesstransformationinst_node.setAttribute('STAGE', '0')
    sesstransformationinst_node.setAttribute('TRANSFORMATIONNAME', s_tbl_name)
    sesstransformationinst_node.setAttribute('TRANSFORMATIONTYPE', 'Source Definition')
    session_node.appendChild(sesstransformationinst_node)

    sesstransformationinst_node = document.createElement('SESSTRANSFORMATIONINST')
    sesstransformationinst_node.setAttribute('ISREPARTITIONPOINT', 'YES')
    sesstransformationinst_node.setAttribute('PARTITIONTYPE', 'PASS THROUGH')
    sesstransformationinst_node.setAttribute('PIPELINE', '1')
    sesstransformationinst_node.setAttribute('SINSTANCENAME', sq_name)
    sesstransformationinst_node.setAttribute('STAGE', '2')
    sesstransformationinst_node.setAttribute('TRANSFORMATIONNAME', sq_name)
    sesstransformationinst_node.setAttribute('TRANSFORMATIONTYPE', 'Source Qualifier')
    session_node.appendChild(sesstransformationinst_node)

    sesstransformationinst_node = document.createElement('SESSTRANSFORMATIONINST')
    sesstransformationinst_node.setAttribute('ISREPARTITIONPOINT', 'NO')
    sesstransformationinst_node.setAttribute('PIPELINE', '1')
    sesstransformationinst_node.setAttribute('SINSTANCENAME', 'EXPTRANS')
    sesstransformationinst_node.setAttribute('STAGE', '2')
    sesstransformationinst_node.setAttribute('TRANSFORMATIONNAME', 'EXPTRANS')
    sesstransformationinst_node.setAttribute('TRANSFORMATIONTYPE', 'Expression')
    partition_node = document.createElement('PARTITION')
    partition_node.setAttribute('DESCRIPTION', '')
    partition_node.setAttribute('NAME', '分区编号1')
    sesstransformationinst_node.appendChild(partition_node)
    session_node.appendChild(sesstransformationinst_node)

    configreference_node = document.createElement('CONFIGREFERENCE')
    configreference_node.setAttribute('REFOBJECTNAME', 'default_session_config')
    configreference_node.setAttribute('TYPE', 'Session config')
    attribute_node = document.createElement('ATTRIBUTE')
    attribute_node.setAttribute('NAME', 'DateTime Format String')
    attribute_node.setAttribute('VALUE', 'YYYY/MM/DD HH24:MI:SS.US')
    configreference_node.appendChild(attribute_node)
    session_node.appendChild(configreference_node)

    sessioncomponent_node = document.createElement('SESSIONCOMPONENT')
    sessioncomponent_node.setAttribute('REFOBJECTNAME', 'post_session_success_command')
    sessioncomponent_node.setAttribute('REUSABLE', 'NO')
    sessioncomponent_node.setAttribute('TYPE', 'Post-session success command')

    task_node = document.createElement('TASK')
    task_node.setAttribute('DESCRIPTION', '')
    task_node.setAttribute('NAME', 'post_session_success_command')
    task_node.setAttribute('REUSABLE', 'NO')
    task_node.setAttribute('TYPE', 'Command')
    task_node.setAttribute('VERSIONNUMBER', '1')

    attribute_node = document.createElement('ATTRIBUTE')
    attribute_node.setAttribute('NAME', 'Fail task if any command fails')
    attribute_node.setAttribute('VALUE', 'NO')
    task_node.appendChild(attribute_node)

    attribute_node = document.createElement('ATTRIBUTE')
    attribute_node.setAttribute('NAME', 'Recovery Strategy')
    attribute_node.setAttribute('VALUE', 'Fail task and continue workflow')
    task_node.appendChild(attribute_node)

    shell_value = 'sh /root/Informatica/mysql_load_GBase8a.sh ' + m_db.lower() + ' ' + m_tbl_name.lower() + ' ' + s_db.lower() + ' ' + \
                  load_sql_query_string.lower() + 'informatica_row_id,informatica_flag,informatica_date_time,ext_ogg_seq,extend_field_src_system,extend_field_valid_flag,extend_field_update_flag,extend_field_update_time'
    shell_value2 = 'sh /root/Informatica/data_qingxi.sh ' + m_db.lower() + ' ' + m_tbl_name.lower()

    valuepair_node = document.createElement('VALUEPAIR')
    valuepair_node.setAttribute('EXECORDER', '1')
    valuepair_node.setAttribute('NAME', '命令1')
    valuepair_node.setAttribute('REVERSEASSIGNMENT', 'NO')
    valuepair_node.setAttribute('VALUE', shell_value)
    task_node.appendChild(valuepair_node)

    valuepair_node = document.createElement('VALUEPAIR')
    valuepair_node.setAttribute('EXECORDER', '2')
    valuepair_node.setAttribute('NAME', '命令2')
    valuepair_node.setAttribute('REVERSEASSIGNMENT', 'NO')
    valuepair_node.setAttribute('VALUE', shell_value2)
    task_node.appendChild(valuepair_node)

    sessioncomponent_node.appendChild(task_node)
    session_node.appendChild(sessioncomponent_node)

    sessioncomponent_node = document.createElement('SESSIONCOMPONENT')
    sessioncomponent_node.setAttribute('REFOBJECTNAME', 'presession_variable_assignment')
    sessioncomponent_node.setAttribute('REUSABLE', 'NO')
    sessioncomponent_node.setAttribute('TYPE', 'Pre-session variable assignment')

    task_node = document.createElement('TASK')
    task_node.setAttribute('DESCRIPTION', '')
    task_node.setAttribute('NAME', 'presession_variable_assignment')
    task_node.setAttribute('REUSABLE', 'NO')
    task_node.setAttribute('TYPE', 'Command')
    task_node.setAttribute('VERSIONNUMBER', '1')

    attribute_node = document.createElement('ATTRIBUTE')
    attribute_node.setAttribute('NAME', 'Fail task if any command fails')
    attribute_node.setAttribute('VALUE', 'NO')
    task_node.appendChild(attribute_node)

    attribute_node = document.createElement('ATTRIBUTE')
    attribute_node.setAttribute('NAME', 'Recovery Strategy')
    attribute_node.setAttribute('VALUE', 'Fail task and continue workflow')
    task_node.appendChild(attribute_node)

    sessioncomponent_node.appendChild(task_node)
    session_node.appendChild(sessioncomponent_node)

    sessioncomponent_node = document.createElement('SESSIONCOMPONENT')
    sessioncomponent_node.setAttribute('REFOBJECTNAME', 'postsession_success_variable_assignment')
    sessioncomponent_node.setAttribute('REUSABLE', 'NO')
    sessioncomponent_node.setAttribute('TYPE', 'Post-session success variable assignment')

    task_node = document.createElement('TASK')
    task_node.setAttribute('DESCRIPTION', '')
    task_node.setAttribute('NAME', 'postsession_success_variable_assignment')
    task_node.setAttribute('REUSABLE', 'NO')
    task_node.setAttribute('TYPE', 'Command')
    task_node.setAttribute('VERSIONNUMBER', '1')

    attribute_node = document.createElement('ATTRIBUTE')
    attribute_node.setAttribute('NAME', 'Fail task if any command fails')
    attribute_node.setAttribute('VALUE', 'NO')
    task_node.appendChild(attribute_node)

    attribute_node = document.createElement('ATTRIBUTE')
    attribute_node.setAttribute('NAME', 'Recovery Strategy')
    attribute_node.setAttribute('VALUE', 'Fail task and continue workflow')
    task_node.appendChild(attribute_node)

    sessioncomponent_node.appendChild(task_node)
    session_node.appendChild(sessioncomponent_node)

    sessioncomponent_node = document.createElement('SESSIONCOMPONENT')
    sessioncomponent_node.setAttribute('REFOBJECTNAME', 'postsession_failure_variable_assignment')
    sessioncomponent_node.setAttribute('REUSABLE', 'NO')
    sessioncomponent_node.setAttribute('TYPE', 'Post-session failure variable assignment')

    task_node = document.createElement('TASK')
    task_node.setAttribute('DESCRIPTION', '')
    task_node.setAttribute('NAME', 'postsession_failure_variable_assignment')
    task_node.setAttribute('REUSABLE', 'NO')
    task_node.setAttribute('TYPE', 'Command')
    task_node.setAttribute('VERSIONNUMBER', '1')

    attribute_node = document.createElement('ATTRIBUTE')
    attribute_node.setAttribute('NAME', 'Fail task if any command fails')
    attribute_node.setAttribute('VALUE', 'NO')
    task_node.appendChild(attribute_node)

    attribute_node = document.createElement('ATTRIBUTE')
    attribute_node.setAttribute('NAME', 'Recovery Strategy')
    attribute_node.setAttribute('VALUE', 'Fail task and continue workflow')
    task_node.appendChild(attribute_node)

    sessioncomponent_node.appendChild(task_node)
    session_node.appendChild(sessioncomponent_node)

    sessionextension_node = document.createElement('SESSIONEXTENSION')
    sessionextension_node.setAttribute('NAME', 'File Writer')
    sessionextension_node.setAttribute('SINSTANCENAME', m_tbl_name)
    sessionextension_node.setAttribute('SUBTYPE', 'File Writer')
    sessionextension_node.setAttribute('TRANSFORMATIONTYPE', 'Target Definition')
    sessionextension_node.setAttribute('TYPE', 'WRITER')

    connectionreference_node = document.createElement('CONNECTIONREFERENCE')
    connectionreference_node.setAttribute('CNXREFNAME', 'Connection')
    connectionreference_node.setAttribute('CONNECTIONNAME', '')
    connectionreference_node.setAttribute('CONNECTIONNUMBER', '1')
    connectionreference_node.setAttribute('CONNECTIONSUBTYPE', '')
    connectionreference_node.setAttribute('CONNECTIONTYPE', '')
    connectionreference_node.setAttribute('VARIABLE', '')
    sessionextension_node.appendChild(connectionreference_node)

    m_tblname_lower = m_tbl_name.lower()
    s_db_lower = s_db.lower()

    sessionextension_attribut_attrs = [ \
        {'NAME': 'Merge Type', 'VALUE': 'No Merge'}, \
        {'NAME': 'Merge File Directory', 'VALUE': '$PMTargetFileDir'}, \
        {'NAME': 'Merge File Name', 'VALUE': m_tblname_lower + '.out'}, \
        {'NAME': 'Append if Exists', 'VALUE': 'NO'}, \
        {'NAME': 'Create Target Directory', 'VALUE': 'NO'}, \
        {'NAME': 'Header Options', 'VALUE': 'No Header'}, \
        {'NAME': 'Header Command', 'VALUE': ''}, \
        {'NAME': 'Footer Command', 'VALUE': ''}, \
        {'NAME': 'Output Type', 'VALUE': 'File'}, \
        {'NAME': 'Merge Command', 'VALUE': ''}, \
        {'NAME': 'Output file directory', 'VALUE': '$PMTargetFileDir'}, \
        {'NAME': 'Output filename', 'VALUE': s_db_lower + '_' + m_tblname_lower + '1.out'}, \
        {'NAME': 'Reject file directory', 'VALUE': '$PMBadFileDir'}, \
        {'NAME': 'Reject filename', 'VALUE': m_tblname_lower + '1.bad'}, \
        {'NAME': 'Command', 'VALUE': ''}, \
        {'NAME': 'Codepage Parameter', 'VALUE': ''} \
        ]
    for sext_arr in sessionextension_attribut_attrs:
        attribute_node = document.createElement('ATTRIBUTE')
        attribute_node.setAttribute('NAME', sext_arr['NAME'])
        attribute_node.setAttribute('VALUE', sext_arr['VALUE'])
        sessionextension_node.appendChild(attribute_node)

    session_node.appendChild(sessionextension_node)

    sessionextension_node = document.createElement('SESSIONEXTENSION')
    sessionextension_node.setAttribute('DSQINSTNAME', sq_name)
    sessionextension_node.setAttribute('DSQINSTTYPE', 'Source Qualifier')
    sessionextension_node.setAttribute('NAME', 'Relational Reader')
    sessionextension_node.setAttribute('SINSTANCENAME', s_tbl_name)
    sessionextension_node.setAttribute('SUBTYPE', 'Relational Reader')
    sessionextension_node.setAttribute('TRANSFORMATIONTYPE', 'Source Definition')
    sessionextension_node.setAttribute('TYPE', 'READER')
    session_node.appendChild(sessionextension_node)

    sessionextension_node = document.createElement('SESSIONEXTENSION')
    sessionextension_node.setAttribute('NAME', 'Relational Reader')
    sessionextension_node.setAttribute('SINSTANCENAME', sq_name)
    sessionextension_node.setAttribute('SUBTYPE', 'Relational Reader')
    sessionextension_node.setAttribute('TRANSFORMATIONTYPE', 'Source Qualifier')
    sessionextension_node.setAttribute('TYPE', 'READER')

    connectionreference_node = document.createElement('CONNECTIONREFERENCE')
    connectionreference_node.setAttribute('CNXREFNAME', 'DB Connection')
    connectionreference_node.setAttribute('CONNECTIONNAME', 'mysql' + s_db)
    connectionreference_node.setAttribute('CONNECTIONNUMBER', '1')
    connectionreference_node.setAttribute('CONNECTIONSUBTYPE', 'ODBC')
    connectionreference_node.setAttribute('CONNECTIONTYPE', 'Relational')
    connectionreference_node.setAttribute('VARIABLE', '')
    sessionextension_node.appendChild(connectionreference_node)

    session_node.appendChild(sessionextension_node)
    session_attribute_arrs = [ \
        {'NAME': 'General Options', 'VALUE': ''}, \
        {'NAME': 'Write Backward Compatible Session Log File', 'VALUE': 'NO'}, \
        {'NAME': 'Session Log File Name', 'VALUE': session_name + '.log'}, \
        {'NAME': 'Session Log File directory', 'VALUE': '$PMSessionLogDir'}, \
        {'NAME': 'Parameter Filename', 'VALUE': ''}, \
        {'NAME': 'Enable Test Load', 'VALUE': 'NO'}, \
        {'NAME': '$Source connection value', 'VALUE': ''}, \
        {'NAME': '$Target connection value', 'VALUE': ''}, \
        {'NAME': 'Treat source rows as', 'VALUE': 'Insert'}, \
        {'NAME': 'Commit Type', 'VALUE': 'Target'}, \
        {'NAME': 'Commit Interval', 'VALUE': '10000'}, \
        {'NAME': 'Commit On End Of File', 'VALUE': 'YES'}, \
        {'NAME': 'Rollback Transactions on Errors', 'VALUE': 'NO'}, \
        {'NAME': 'Recovery Strategy', 'VALUE': 'Fail task and continue workflow'}, \
        {'NAME': 'Java Classpath', 'VALUE': ''}, \
        {'NAME': 'Performance', 'VALUE': ''}, \
        {'NAME': 'DTM buffer size', 'VALUE': '24000000'}, \
        {'NAME': 'Collect performance data', 'VALUE': 'NO'}, \
        {'NAME': 'Write performance data to repository', 'VALUE': 'NO'}, \
        {'NAME': 'Incremental Aggregation', 'VALUE': 'NO'}, \
        {'NAME': 'Session retry on deadlock', 'VALUE': 'NO'}, \
        {'NAME': 'Pushdown Optimization', 'VALUE': 'None'}, \
        {'NAME': 'Allow Temporary View for Pushdown', 'VALUE': 'NO'}, \
        {'NAME': 'Allow Temporary Sequence for Pushdown', 'VALUE': 'NO'}, \
        {'NAME': 'Allow Pushdown for User Incompatible Connections', 'VALUE': 'NO'} \
        ]
    for s_att in session_attribute_arrs:
        session_attribute_node = document.createElement('ATTRIBUTE')
        session_attribute_node.setAttribute('NAME', s_att['NAME'])
        session_attribute_node.setAttribute('VALUE', s_att['VALUE'])
        session_node.appendChild(session_attribute_node)

    workflow_node.appendChild(session_node)

    taskinstance_node = document.createElement('TASKINSTANCE')
    taskinstance_node.setAttribute('DESCRIPTION', '')
    taskinstance_node.setAttribute('ISENABLED', 'YES')
    taskinstance_node.setAttribute('NAME', '启动')
    taskinstance_node.setAttribute('REUSABLE', 'NO')
    taskinstance_node.setAttribute('TASKNAME', '启动')
    taskinstance_node.setAttribute('TASKTYPE', 'Start')
    workflow_node.appendChild(taskinstance_node)

    taskinstance_node = document.createElement('TASKINSTANCE')
    taskinstance_node.setAttribute('DESCRIPTION', '')
    taskinstance_node.setAttribute('FAIL_PARENT_IF_INSTANCE_DID_NOT_RUN', 'NO')
    taskinstance_node.setAttribute('FAIL_PARENT_IF_INSTANCE_FAILS', 'YES')
    taskinstance_node.setAttribute('ISENABLED', 'YES')
    taskinstance_node.setAttribute('NAME', session_name)
    taskinstance_node.setAttribute('REUSABLE', 'NO')
    taskinstance_node.setAttribute('TASKNAME', session_name)
    taskinstance_node.setAttribute('TASKTYPE', 'Session')
    taskinstance_node.setAttribute('TREAT_INPUTLINK_AS_AND', 'YES')

    workflow_node.appendChild(taskinstance_node)

    workflowlink_node = document.createElement('WORKFLOWLINK')
    workflowlink_node.setAttribute('CONDITION', '')
    workflowlink_node.setAttribute('FROMTASK', '启动')
    workflowlink_node.setAttribute('TOTASK', session_name)
    workflow_node.appendChild(workflowlink_node)

    wfvariable_arrs = [ \
        {'DATATYPE': 'date/time', 'DESCRIPTION': 'The time this task started', 'NAME': '$启动.StartTime'}, \
        {'DATATYPE': 'date/time', 'DESCRIPTION': 'The time this task completed', 'NAME': '$启动.EndTime'}, \
        {'DATATYPE': 'integer', 'DESCRIPTION': 'Status of this task\'s execution', 'NAME': '$启动.Status'}, \
        {'DATATYPE': 'integer', 'DESCRIPTION': 'Status of the previous task that is not disabled',
         'NAME': '$启动.PrevTaskStatus'}, \
        {'DATATYPE': 'integer', 'DESCRIPTION': 'Error code for this task\'s execution', 'NAME': '$启动.ErrorCode'}, \
        {'DATATYPE': 'string', 'DESCRIPTION': 'Error message for this task\'s execution', 'NAME': '$启动.ErrorMsg'}, \
        {'DATATYPE': 'date/time', 'DESCRIPTION': 'The time this task started',
         'NAME': '$' + session_name + '.StartTime'}, \
        {'DATATYPE': 'date/time', 'DESCRIPTION': 'The time this task completed',
         'NAME': '$' + session_name + '.EndTime'}, \
        {'DATATYPE': 'integer', 'DESCRIPTION': 'Status of this task\'s execution',
         'NAME': '$' + session_name + '.Status'}, \
        {'DATATYPE': 'integer', 'DESCRIPTION': 'Status of the previous task that is not disabled',
         'NAME': '$' + session_name + '.PrevTaskStatus'}, \
        {'DATATYPE': 'integer', 'DESCRIPTION': 'Error code for this task\'s execution',
         'NAME': '$' + session_name + '.ErrorCode'}, \
        {'DATATYPE': 'string', 'DESCRIPTION': 'Error message for this task\'s execution',
         'NAME': '$' + session_name + '.ErrorMsg'}, \
        {'DATATYPE': 'integer', 'DESCRIPTION': 'Rows successfully read',
         'NAME': '$' + session_name + '.SrcSuccessRows'}, \
        {'DATATYPE': 'integer', 'DESCRIPTION': 'Rows failed to read', 'NAME': '$' + session_name + '.SrcFailedRows'}, \
        {'DATATYPE': 'integer', 'DESCRIPTION': 'Rows successfully loaded',
         'NAME': '$' + session_name + '.TgtSuccessRows'}, \
        {'DATATYPE': 'integer', 'DESCRIPTION': 'Rows failed to load', 'NAME': '$' + session_name + '.TgtFailedRows'}, \
        {'DATATYPE': 'integer', 'DESCRIPTION': 'Total number of transformation errors',
         'NAME': '$' + session_name + '.TotalTransErrors'}, \
        {'DATATYPE': 'integer', 'DESCRIPTION': 'First error code', 'NAME': '$' + session_name + '.FirstErrorCode'}, \
        {'DATATYPE': 'string', 'DESCRIPTION': 'First error message', 'NAME': '$' + session_name + '.FirstErrorMsg'} \
        ]

    for wfvb in wfvariable_arrs:
        workflowvariable_node = document.createElement('WORKFLOWVARIABLE')
        workflowvariable_node.setAttribute('DATATYPE', wfvb['DATATYPE'])
        workflowvariable_node.setAttribute('DEFAULTVALUE', '')
        workflowvariable_node.setAttribute('DESCRIPTION', wfvb['DESCRIPTION'])
        workflowvariable_node.setAttribute('ISNULL', 'NO')
        workflowvariable_node.setAttribute('ISPERSISTENT', 'NO')
        workflowvariable_node.setAttribute('NAME', wfvb['NAME'])
        workflowvariable_node.setAttribute('USERDEFINED', 'NO')

        workflow_node.appendChild(workflowvariable_node)

    attribute_arrs = [
        {'NAME': 'Parameter Filename', 'VALUE': ''}, \
        {'NAME': 'Write Backward Compatible Workflow Log File', 'VALUE': 'NO'}, \
        {'NAME': 'Workflow Log File Name', 'VALUE': 'wf_m_1_DWD_ICMS_PM_ACCODEMTSITUATION_WEEK.log'}, \
        {'NAME': 'Workflow Log File Directory', 'VALUE': '$PMWorkflowLogDir'}, \
        {'NAME': 'Save Workflow log by', 'VALUE': 'By runs'}, \
        {'NAME': 'Save workflow log for these runs', 'VALUE': '0'}, \
        {'NAME': 'Service Name', 'VALUE': ''}, \
        {'NAME': 'Service Timeout', 'VALUE': '0'}, \
        {'NAME': 'Is Service Visible', 'VALUE': 'NO'}, \
        {'NAME': 'Is Service Protected', 'VALUE': 'NO'}, \
        {'NAME': 'Enable HA recovery', 'VALUE': 'NO'}, \
        {'NAME': 'Automatically recover terminated tasks', 'VALUE': 'NO'}, \
        {'NAME': 'Service Level Name', 'VALUE': 'Default'}, \
        {'NAME': 'Allow concurrent run with unique run instance name', 'VALUE': 'NO'}, \
        {'NAME': 'Allow concurrent run with same run instance name', 'VALUE': 'NO'}, \
        {'NAME': 'Maximum number of concurrent runs', 'VALUE': '0'}, \
        {'NAME': 'Assigned Web Services Hubs', 'VALUE': ''}, \
        {'NAME': 'Maximum number of concurrent runs per Hub', 'VALUE': '1000'}, \
        {'NAME': 'Expected Service Time', 'VALUE': '1'} \
        ]
    for attr in attribute_arrs:
        workflow_attribute_node = document.createElement('ATTRIBUTE')
        workflow_attribute_node.setAttribute('NAME', attr['NAME'])
        workflow_attribute_node.setAttribute('VALUE', attr['VALUE'])
        workflow_node.appendChild(workflow_attribute_node)

    folder_node.appendChild(workflow_node)


def listfiles(path):
    files = []
    filenum = 0
    # 获取路径下的所有文件和文件夹 的名字
    list = os.listdir(path)
    for line in list:
        filepath = os.path.join(path, line)
        if os.path.isfile(filepath):
            files.append(line)
    return files


def main(excel_file, xml_file):
    # 读取模板xml文件
    s_filename = r'C:\Users\mengx\INFA_XML_project\mode.xml'
    print(excel_file)
    doc = Dom.parse(s_filename)
    inf_dic = read_xlsx(excel_file)
    m_table_names = set()
    for key, value in inf_dic.items():
        m_table_names.add(key)

    tbl_count = 0
    for m_table_name in m_table_names:
        if m_table_name == '':
            continue
        tbl_count = tbl_count + 1
        s_table_name = inf_dic[m_table_name][0][2]
        # 源库名，用于source标签中的DBDNAME属性
        s_dbname = inf_dic[m_table_name][0][5]
        # 目标库名
        m_dbname = inf_dic[m_table_name][0][6]

        contents_xml = Dom.parse(r'c:\Users\mengx\INFA_XML_project\contents.xml')
        config_node = contents_xml.getElementsByTagName('CONFIG')

        create_tag(doc, s_table_name, s_dbname, inf_dic, m_table_name, tbl_count, m_dbname, config_node[0])

    f = open(xml_file + '.xml',
             'w', encoding='utf-8')
    doc.toprettyxml(encoding='utf-8')
    doc.writexml(f)
    f.close()


if __name__ == '__main__':
    source_path = r'C:\Users\mengx\INFA_XML_project\source'
    result_path = r'C:\Users\mengx\INFA_XML_project\result'
    fls = listfiles(source_path)
    for filename in fls:
        source_filename = source_path + '\\' + filename
        result_filename = result_path + '\\' + filename
        sys_names = re.findall(r'(.*?)[_.](.*)', filename)
        system_name = sys_names[0][0]
        print('正在处理{0} 系统。。。'.format(system_name))
        main(source_filename, result_filename)
