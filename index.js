/**
 * export excel file for sqlite db
 */
var sqlite3 = require('sqlite3').verbose(),
    db = new sqlite3.Database('./kmcrj.db'),
    excel = require('exceljs'),
    workbook_product,
    workbook_unit,
    options_product = {
        filename: './商品信息.xlsx',
        useStyles: true,
        useSharedStrings: true
    },
    options_unit = {
        filename: '往来单位信息.xlsx',
        useStyles: true,
        useSharedStrings: true
    };
workbook_product = new excel.stream.xlsx.WorkbookWriter(options_product);
workbook_unit = new excel.stream.xlsx.WorkbookWriter(options_unit);
let worksheet_product = workbook_product.addWorksheet('商品信息'),
    worksheet_unit = workbook_unit.addWorksheet('往来单位信息');
worksheet_product.columns = [{
    header: '商品编码',
    key: 'productsn',
    width: 10
},
{
    header: '商品名称',
    key: 'productname',
    width: 25
},
{
    header: '成本算法',
    key: 'mode',
    width: 10,
},
{
    header: '条码',
    key: 'barcode',
    width: 20,
},
{
    header: '基本单位',
    key: 'productunit',
    width: 10,
},
{
    header: '规格',
    key: 'productmodel',
    width: 10,
},
{
    header: '型号',
    key: 'productmodel1',
    width: 10,
},
{
    header: '产地',
    key: 'productplace',
    width: 30,
},
{
    header: '预设售价1',
    key: 'budget1',
    width: 10,
},
{
    header: '预设售价2',
    key: 'budget2',
    width: 10,
},
{
    header: '预设售价3',
    key: 'budget3',
    width: 10,
},
{
    header: '预设售价4',
    key: 'budget4',
    width: 10,
},
{
    header: '预设售价5',
    key: 'budget5',
    width: 10,
},
{
    header: '零售价',
    key: 'retailprice',
    width: 10,
},
{
    header: '会员价',
    key: 'vipprice',
    width: 10,
},
{
    header: '提前报警天数',
    key: 'warningDays',
    width: 10,
},
{
    header: '停用标志',
    key: 'productstatus',
    width: 10,
},
{
    header: '是否换货',
    key: 'exchange',
    width: 10,
},
{
    header: '最近进价',
    key: 'lastinprice',
    width: 10,
},
{
    header: '最近售价',
    key: 'lastoutprice',
    width: 10,
},
{
    header: '最低售价',
    key: 'lowoutprice',
    width: 10,
},
{
    header: '自定义1编码',
    key: 'englishname1',
    width: 15,
},
{
    header: '自定义2编码',
    key: 'englishname2',
    width: 15,
},
{
    header: '自定义3编码',
    key: 'englishname3',
    width: 15,
},
{
    header: '商品别名',
    key: 'abbr',
    width: 25,
},
{
    header: '助记码',
    key: 'initials',
    width: 20,
},
{
    header: '换算率1',
    key: 'exchangerate1',
    width: 10,
},
{
    header: '辅助单位1',
    key: 'productunit1',
    width: 10,
},
{
    header: '辅助单位2',
    key: 'productunit2',
    width: 10,
},
{
    header: '换算率2',
    key: 'exchangerate2',
    width: 10,
},
{
    header: '保质期（天）',
    key: 'quality',
    width: 15,
},
{
    header: '备注',
    key: 'comment',
    width: 20,
},
{
    header: '替代品',
    key: 'replace',
    width: 10,
},
];
worksheet_unit.columns = [{
    header: '单位编码',
    key: "partysn",
    width: 10
},
{
    header: '单位名称',
    key: "partyname",
    width: 25
},
{
    header: '类别',
    key: "partyclass",
    width: 10,
},
{
    header: '应收应付款账期',
    key: "ysyfdate",
    width: 20,
},
{
    header: '退换货期限（天）',
    key: "thdate",
    width: 10,
},
{
    header: '联系人',
    key: "contactname",
    width: 10,
},
{
    header: '联系电话',
    key: "cellphone",
    width: 15,
},
{
    header: '传真电话',
    key: "fox",
    width: 15,
},
{
    header: '地区',
    key: "region",
    width: 10,
},
{
    header: '联系地址',
    key: "deliveryaddressas",
    width: 10,
},
{
    header: '单位别名',
    key: "abbr",
    width: 25,
},
{
    header: '助记码',
    key: "initials",
    width: 20,
},
{
    header: '移动电话',
    key: "telphone",
    width: 15,
},
{
    header: '电子邮件',
    key: "email",
    width: 20,
},
{
    header: '邮政编号',
    key: "postcode",
    width: 10,
},
{
    header: '银行账号',
    key: "accountno",
    width: 20,
},
{
    header: '税号',
    key: "taxid",
    width: 20,
},
{
    header: '信用额度',
    key: "creditlimit",
    width: 10,
},
{
    header: '保证金',
    key: "cautionmoney",
    width: 10,
},
{
    header: '最近退换货比率',
    key: "thrate",
    width: 10,
},
{
    header: '期初应付',
    key: "payables",
    width: 10,
},
{
    header: '期初应收',
    key: "receivables",
    width: 10,
},
{
    header: '预设售价级别(入库)',
    key: "presetsalesin",
    width: 10,
},
{
    header: '预设售价比例(入库)',
    key: "presetsalesratein",
    width: 10,
},
{
    header: '预设售价级别(出库)',
    key: "presetsalesout",
    width: 10,
},
{
    header: '预设售价比例(出库)',
    key: "presetsalesrateout",
    width: 10,
},
{
    header: '备注',
    key: "comment",
    width: 20,
},
{
    header: '财务通供应商编码',
    key: "cwtsupplierno",
    width: 10,
},
{
    header: '财务通供应商名称',
    key: "cwtsuppliername",
    width: 20,
},
{
    header: '财务通客户编码',
    key: "cwtclientno",
    width: 10,
},
{
    header: '财务通客户商名称',
    key: "cwtclientname",
    width: 20,
}
];
db.serialize(function () {
    //商品信息
    let sql_product = 'select productsn, productname, mode, ifnull(json_extract(barcodelist, "$[0].barcode"), brand) as "barcode", json_extract(priceconfig,"$.measure.productunit") as "productunit", productmodel, productmodel as "productmodel1", null as "productplace", json_extract(priceconfig,"$.prices[0].price1") as "budget1", json_extract(priceconfig,"$.prices[0].price2") as "budget2", json_extract(priceconfig,"$.prices[0].price3") as "budget3", json_extract(priceconfig,"$.prices[0].price4") as "budget4", json_extract(priceconfig,"$.prices[0].price5") as "budget5", json_extract(priceconfig,"$.prices[0].retailprice") as "retailprice", null as "vipprice",  warningDays, productstatus, null as "exchange", json_extract(priceconfig,"$.prices.lastinprice") as "lastinprice", json_extract(priceconfig,"$.prices.lastoutprice") as "lastoutprice", json_extract(priceconfig,"$.prices[0].lowoutprice") as "lowoutprice", json_extract(customized, "$[0].englishname") as "englishname1",json_extract(customized, "$[1].englishname") as "englishname2", json_extract(customized, "$[2].englishname") as "englishname3", abbr, initials, json_extract(priceconfig,"$.prices[1].exchangerate") as "exchangerate1", json_extract(priceconfig,"$.prices[1].productunit") as "productunit1", json_extract(priceconfig,"$.prices[2].productunit") as "productunit2", json_extract(priceconfig,"$.prices[2].exchangerate") as "exchangerate2",  null as "quality", comment, null as "replace" from dbproducts,(select json_extract(v, "$.warningDays") as warningDays from dbsystemconfig where k = "warning_config")',
        sql_unit = 'select partysn, partyname, partyclass, json_extract(clearingform,"$.gathering.gysyfdate ")||"/"||json_extract(clearingform,"$.payment.pysyfdate") as "ysyfdate", null as "thdate",json_extract(contact,"$.name") as "contactname", json_extract(contact,"$.cellphone") "cellphone", json_extract(contact,"$.fox") as "fox", null as "region", json_extract(contact,"$.deliveryaddressas") "deliveryaddressas", abbr, initials, json_extract(contact,"$.telephone") as "telphone",  json_extract(contact,"$.email") as "email", json_extract(contact,"$.postcode") as "postcode", json_extract(partyinfo,"$.accountno") as "accountno", json_extract(partyinfo,"$.taxid") as "taxid", json_extract(taxinfo,"$.creditlimit") as "creditlimit",null as "cautionmoney",null as  "thrate", payables,receivables, null as "presetsalesin", null as "presetsalesratein", null as "presetsalesout", null as "presetsalesrateout", comment, null as "cwtsupplierno", null as "cwtsuppliername", null as "cwtclientno", null as "cwtclientname" from dbrelatedparties r left outer join (select partyid, sum(payables) payables, sum(receivables) receivables from dbbills group by partyid) b on r.partyid = b.partyid';
    db.all(sql_product, function (err, rows) {
        if (!!err) throw err;
        for (value of rows) {
            worksheet_product.addRow(value);
        }
        worksheet_product.commit();
        workbook_product.commit()
            .then(function () {
                console.info('商品信息导入完成！')
            });
    });
    db.all(sql_unit, function (err, rows) {
        for (value of rows) {
            worksheet_unit.addRow(value);
        }
        worksheet_unit.commit();
        workbook_unit.commit()
            .then(function () {
                console.info('往来单位信息导入完成！')
            });
    });
});
db.close();



