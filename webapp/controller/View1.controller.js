sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/m/MessageToast",
    "sap/ui/model/json/JSONModel",
    "sap/ui/core/BusyIndicator",
"sap/ui/export/Spreadsheet",
 "sap/ui/export/library"
], function (Controller, MessageToast, JSONModel, BusyIndicator, Spreadsheet, library) {
    "use strict";

    return Controller.extend("uploadfile.controller.View1", {

        onInit: function () {
            this._excelData = [];
        },

        onFileUpload: function (oEvent) {
            const file = oEvent.getParameter("files")[0];
            if (!file) {
                MessageToast.show("No file selected!");
                return;
            }

            BusyIndicator.show(0);  // 🔥 Show spinning loader

            const reader = new FileReader();

            reader.onload = (e) => {
                const data = e.target.result;

                try {
                    // Parse XLSX
                    const workbook = XLSX.read(data, { type: "binary" });
                    const sheetName = workbook.SheetNames[0];
                    const sheet = workbook.Sheets[sheetName];

                    this._excelData = XLSX.utils.sheet_to_json(sheet);

                    MessageToast.show("File loaded successfully!");
                } catch (err) {
                    MessageToast.show("Error reading Excel file!");
                }

                BusyIndicator.hide(); // 🔥 Hide loader
            };

            reader.readAsBinaryString(file);
        },

        handleUploadPress: function () {

            if (!this._excelData.length) {
                MessageToast.show("Upload an Excel file first!");
                return;
            }

            BusyIndicator.show(0); // 🔥 Loader ON

            setTimeout(() => { 
                const oModel = new JSONModel({
                    SalesOrders: this._excelData
                });

                const oTable = this.getView().byId("salesTable");
                oTable.setModel(oModel);

                const oTemplate = oTable.getItems()[0].clone();

                oTable.bindItems({
                    path: "/SalesOrders",
                    template: oTemplate
                });

                oTable.setVisible(true);

                MessageToast.show("Excel data loaded into table!");

                BusyIndicator.hide();  // 🔥 Loader OFF

            }, 500);
        },

        onDownloadSelectedPDF: function () {
            const oTable = this.getView().byId("salesTable");
            const aSelectedItems = oTable.getSelectedItems();

            if (aSelectedItems.length === 0) {
                sap.m.MessageToast.show("Please select at least one row!");
                return;
            }

            const selectedData = aSelectedItems.map(item =>
                item.getBindingContext().getObject()
            );

            const { jsPDF } = window.jspdf;
            const doc = new jsPDF();

            doc.setFontSize(16);
            doc.text("Selected Sales Order Records", 14, 15);

            const columns = [
                { header: "Sales Order ID", dataKey: "SalesOrderID" },
                { header: "Customer Name", dataKey: "SalesOrderName" },
                { header: "Order Date", dataKey: "SalesOrderDate" },
                { header: "Amount", dataKey: "Amount" },
                { header: "Status", dataKey: "Status" }
            ];

            doc.autoTable({
                startY: 25,
                head: [columns.map(col => col.header)],
                body: selectedData.map(item =>
                    columns.map(col => item[col.dataKey] || "")
                ),
                styles: { fontSize: 10 }
            });

            doc.save("SelectedRows.pdf");
        }, onDownloadExcel: function () {
            const oTable = this.getView().byId("salesTable");
            const oBinding = oTable.getBinding("items");

            if (!oBinding) {
                MessageToast.show("No data in table!");
                return;
            }

            const aRows = oBinding.getContexts().map(ctx => ctx.getObject());

            const aCols = [
                { label: "Sales Order ID", property: "SalesOrderID" },
                { label: "Customer Name", property: "SalesOrderName" },
                { label: "Order Date", property: "SalesOrderDate" },
                { label: "Amount", property: "Amount" },
                { label: "Status", property: "Status" }
            ];

            const oSettings = {
                workbook: {
                    columns: aCols
                },
                dataSource: aRows,
                fileName: "SalesOrders.xlsx"
            };

            const oSheet = new Spreadsheet(oSettings);
            oSheet.build()
                .then(() => MessageToast.show("Excel downloaded!"))
                .finally(() => oSheet.destroy());
        }


    });
});

