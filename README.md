export_excel
===
### 使用
> 查询 `sqlite` 数据库表中的数据，并导出到 `excel` 文档中，本案例生成两个 `excel` 文档的 `demo`
```shell
$ npm install    //安装
$ node index.js    //执行demo
or
$ npm start
```
### 依赖
```shell
"exceljs": "^0.7.1",
"sqlite3": "^3.1.13"
```
## 功能代码
```javascript
db.all(sql_product, function (err, rows) {
        if (!!error) throw error;
        for (value of rows) {
            worksheet_product.addRow(value);    //添加行数据
        }
        worksheet_product.commit();
        workbook_product.commit()
            .then(function () {
                console.info('商品信息导入完成！')
            });
    });
```