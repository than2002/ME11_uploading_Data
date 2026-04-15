sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/m/MessageToast",
    "sap/m/MessageBox",
    "sap/ui/model/json/JSONModel",
    "sap/ui/core/BusyIndicator",
    "sap/ui/model/Filter",
    "sap/ui/model/FilterOperator"
], function (Controller, MessageToast, MessageBox, JSONModel, BusyIndicator, Filter, FilterOperator) {
    "use strict";

    return Controller.extend("me11uploadingdata.controller.View1", {

        onInit: function () {
            this._file = null;

            this.getView().setModel(new JSONModel({
                rows: []
            }), "excel");

            // Batch mode – to uploading best mode
           var oModel = this.getOwnerComponent().getModel();

            if (oModel) {
                oModel.setUseBatch(true);
                oModel.setDeferredGroups(["upload"]);
            }
            this._loadXLSXLibrary();
        },

        
        // Excel Handling                                       
       

        _loadXLSXLibrary: function () {
            if (typeof XLSX !== "undefined") {
                return;
            }

            var script = document.createElement("script");
            script.src = sap.ui.require.toUrl("me11uploadingdata/libs/xlsx.full.min.js");
            script.onerror = function () {
                MessageBox.error("XLSX library load failed.");
            };
            document.head.appendChild(script);
        },

        onFileChange: function (oEvent) {
            this._file = oEvent.getParameter("files")[0];
            if (this._file) {
                MessageToast.show("File Selected: " + this._file.name);
            }
        },

        onReadExcel: function () {
            if (!this._file) {
                MessageBox.warning("Please select Excel file");
                return;
            }

            if (typeof XLSX === "undefined") {
                MessageBox.error("XLSX not loaded yet");
                return;
            }

            BusyIndicator.show(0);

            var that = this;
            var reader = new FileReader();

            reader.onload = function (e) {
                try {
                    var workbook = XLSX.read(new Uint8Array(e.target.result), { type: "array" });
                    var sheet = workbook.Sheets[workbook.SheetNames[0]];
                    var rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

                    var aRows = [];
                    var oDupMap = {};

                    rawData.slice(1).forEach(function (row, idx) {

                        var oRow = {
                            RowNo: idx + 1,
                            Lifnr: that._s(row[0]),
                            Matnr: that._s(row[1]),
                            Ekorg: that._s(row[2]),
                            Werks: that._s(row[3]),
                            Normb: that._s(row[4]),
                            Meins: that._s(row[5]),
                            Aplfz: that._s(row[6]),
                            Norbm: that._d(row[7]),
                            Mwskz: that._s(row[8]),
                            Netpr: that._d(row[9]),
                            Peinh: that._i(row[10]),
                            Bprme: that._s(row[11]),
                            Datab: that._date(row[12]),
                            Datbi: that._date(row[13]),
                            Kschl_02: that._s(row[14]),
                            Kbetr_02: that._s(row[15]),
                            Kschl_03: that._s(row[16]),
                            Kbetr_03: that._s(row[17]),
                            StatusText: "",
                            StatusState: "",
                            ValidationMessage: ""
                        };

                        if (that._isBlank(oRow)) {
                            return;
                        }

                        var aErrors = that._validate(oRow);

                        var sKey = [
                            oRow.Lifnr, oRow.Matnr,
                            oRow.Ekorg, oRow.Werks,
                            oRow.Datab, oRow.Datbi
                        ].join("|");

                        if (oDupMap[sKey]) {
                            aErrors.push("Duplicate row in Excel");
                        }
                        oDupMap[sKey] = true;

                        if (aErrors.length) {
                            oRow.StatusText = "Invalid";
                            oRow.StatusState = "Error";
                            oRow.ValidationMessage = aErrors.join(", ");
                        } else {
                            oRow.StatusText = "Ready";
                            oRow.StatusState = "Success";
                            oRow.ValidationMessage = "OK";
                        }

                        aRows.push(oRow);
                    });

                    that.getView().getModel("excel").setProperty("/rows", aRows);
                    that._updateCounts();
                    MessageToast.show("Excel loaded");

                } catch (e) {
                    MessageBox.error("Excel read error");
                } finally {
                    BusyIndicator.hide();
                }
            };

            reader.readAsArrayBuffer(this._file);
        },

        // Searching bar logic
        onSearch: function (oEvent) {
            var sValue = oEvent.getParameter("newValue");
            var oBinding = this.byId("previewTable").getBinding("rows");

            if (!oBinding) {
                return;
            }

            if (!sValue) {
                oBinding.filter([]);
                return;
            }

            var oFilter = new Filter({
                filters: [
                    new Filter("Lifnr", FilterOperator.Contains, sValue),
                    new Filter("Matnr", FilterOperator.Contains, sValue),
                    new Filter("Ekorg", FilterOperator.Contains, sValue),
                    new Filter("Werks", FilterOperator.Contains, sValue),
                    new Filter("StatusText", FilterOperator.Contains, sValue),
                    new Filter("ValidationMessage", FilterOperator.Contains, sValue)
                ],
                and: false
            });

            oBinding.filter(oFilter);
        },


       
        // Send to Backend                                     
        
        onSendToBackend: function () {
            this._revalidate();

            var aValid = this.getView().getModel("excel").getProperty("/rows")
                .filter(r => r.StatusText === "Ready");

            if (!aValid.length) {
                MessageBox.warning("No valid rows to send");
                return;
            }

            this._send(aValid);
        },

        _send: function (aRows) {
            var that = this;
            var oModel = this.getView().getModel();

            var sUploadId = "UPL" + Date.now() + "_" + Math.floor(Math.random() * 1000);

            var oPayload = {
                UploadId: sUploadId,
                ToItems: aRows.map(function (r) {
                    return {
                        UploadId: sUploadId,
                        Lifnr: r.Lifnr,
                        Matnr: r.Matnr,
                        Ekorg: r.Ekorg,
                        Werks: r.Werks,
                        Normb: r.Normb,
                        Meins: r.Meins,
                        Aplfz: r.Aplfz,
                        Norbm: that._fmt(r.Norbm, 3),
                        Mwskz: r.Mwskz,
                        Netpr: that._fmt(r.Netpr, 2),
                        Peinh: String(parseInt(r.Peinh, 10)),
                        Bprme: r.Bprme,
                        Datab: r.Datab,
                        Datbi: r.Datbi,
                        Kschl_02: r.Kschl_02,
                        Kbetr_02: r.Kbetr_02,
                        Kschl_03: r.Kschl_03,
                        Kbetr_03: r.Kbetr_03
                    };
                })
            };

            BusyIndicator.show(0);

            oModel.create("/UploadHeaderSet", oPayload, {
                groupId: "upload",
                success: function (oData) {
                    BusyIndicator.hide();
                    that._applyBackendResult(oData);
                    MessageBox.success("Upload completed");
                },
                error: function (oErr) {
                    BusyIndicator.hide();
                    MessageBox.error("Backend error during upload");
                }
            });

            oModel.submitChanges({ groupId: "upload" });
        },

       
        // Backend Result Mapping                   
       

        _applyBackendResult: function (oData) {
            if (!oData || !oData.ToItems || !oData.ToItems.results) {
                return;
            }

            var aBackend = oData.ToItems.results;
            var aRows = this.getView().getModel("excel").getProperty("/rows");

            aRows.forEach(function (r) {
                var b = aBackend.find(x =>
                    x.Lifnr === r.Lifnr && x.Matnr === r.Matnr
                );

                if (b) {
                    r.StatusText = b.Status === "S" ? "Success" : "Error";
                    r.StatusState = b.Status === "S" ? "Success" : "Error";
                    r.ValidationMessage = b.Message;
                }
            });

            this.getView().getModel("excel").refresh(true);
            this._updateCounts();
        },

       
        // Validation                                        
       

        _validate: function (r) {
            var e = [];
            if (!r.Lifnr) e.push("LIFNR missing");
            if (!r.Matnr) e.push("MATNR missing");
            if (!r.Ekorg) e.push("EKORG missing");
            if (!r.Werks) e.push("WERKS missing");
            if (!r.Meins) e.push("MEINS missing");
            if (!r.Mwskz) e.push("MWSKZ missing");
            if (!r.Norbm || r.Norbm <= 0) e.push("NORBM > 0 required");
            if (!r.Peinh || r.Peinh <= 0) e.push("PEINH > 0 required");
            if (!/^\d{8}$/.test(r.Datab)) e.push("DATAB invalid");
            if (!/^\d{8}$/.test(r.Datbi)) e.push("DATBI invalid");
            if (r.Datab > r.Datbi) e.push("DATAB > DATBI");
            if (r.Kschl_02 && !r.Kbetr_02) e.push("KBETR_02 required");
            if (r.Kschl_03 && !r.Kbetr_03) e.push("KBETR_03 required");
            return e;
        },

        _revalidate: function () {
            var aRows = this.getView().getModel("excel").getProperty("/rows");
            aRows.forEach(r => {
                var e = this._validate(r);
                r.StatusText = e.length ? "Invalid" : "Ready";
                r.StatusState = e.length ? "Error" : "Success";
                r.ValidationMessage = e.join(", ") || "OK";
            });
            this.getView().getModel("excel").refresh(true);
        },

       
        // Helpers
       

        _s: v => v ? String(v).trim() : "",
        _d: v => v ? String(v).replace(/,/g, "") : "",
        _i: v => v ? parseInt(v, 10) : "",
        _fmt: (v, s) => parseFloat(v || 0).toFixed(s),

       _date: function (v) {
            if (!v) {
                return "";
            }

            //  Excel serial date
            if (typeof v === "number") {
                var d1 = new Date(Math.round((v - 25569) * 86400 * 1000));
                return this._yyyyMMdd(d1);
            }

            var s = String(v).trim();

            //  Already YYYYMMDD
            if (/^\d{8}$/.test(s)) {
                return s;
            }

            //  DD.MM.YYYY | DD-MM-YYYY | DD/MM/YYYY
            var sep = s.includes(".") ? "." : s.includes("-") ? "-" : s.includes("/") ? "/" : null;
            if (sep) {
                var p = s.split(sep);
                if (p.length === 3) {
                    var d2 = new Date(p[2], p[1] - 1, p[0]);
                    if (!isNaN(d2.getTime())) {
                        return this._yyyyMMdd(d2);
                    }
                }
            }

            //  Fallback : any valid date string
            var d3 = new Date(s);
            if (!isNaN(d3.getTime())) {
                return this._yyyyMMdd(d3);
            }

            return "";
        },
        _yyyyMMdd: function (d) {
            return d.getFullYear() +
                String(d.getMonth() + 1).padStart(2, "0") +
                String(d.getDate()).padStart(2, "0");
        },
        _isBlank: function (r) {
            return !(r.Lifnr || r.Matnr || r.Ekorg || r.Werks);
        },

        _updateCounts: function () {
            var a = this.getView().getModel("excel").getProperty("/rows");
            this.byId("txtRows").setText("Rows: " + a.length);
            this.byId("txtValid").setText("Valid: " + a.filter(x => x.StatusText === "Ready").length);
            this.byId("txtInvalid").setText("Invalid: " + a.filter(x => x.StatusText === "Invalid").length);
        },

        onDownloadExcel: function () {
            var aRows = this.getView().getModel("excel").getProperty("/rows");

            if (!aRows || !aRows.length) {
                sap.m.MessageBox.warning("No data to download");
                return;
            }

            if (typeof XLSX === "undefined") {
                sap.m.MessageBox.error("XLSX library not loaded");
                return;
            }

            var aExcelData = [[
                "Supplier","Material","Purch Org","Plant","Normb","UOM","Delivery Time",
                "PO Qty","Tax Code","Net Price","Price Unit","Order Price Unit",
                "Start Date","End Date",
                "Cond Type 2","Cond Value 2",
                "Cond Type 3","Cond Value 3",
                "Status","Message"
            ]];

            aRows.forEach(function (r) {
                aExcelData.push([
                    r.Lifnr, r.Matnr, r.Ekorg, r.Werks, r.Normb,
                    r.Meins, r.Aplfz, r.Norbm, r.Mwskz, r.Netpr,
                    r.Peinh, r.Bprme, r.Datab, r.Datbi,
                    r.Kschl_02, r.Kbetr_02,
                    r.Kschl_03, r.Kbetr_03,
                    r.StatusText, r.ValidationMessage
                ]);
            });

            var wb = XLSX.utils.book_new();
            var ws = XLSX.utils.aoa_to_sheet(aExcelData);
            XLSX.utils.book_append_sheet(wb, ws, "Upload_Result");

            XLSX.writeFile(wb, "ME11_Upload_Result.xlsx");
        },
        onClear: function () {
            this.getView().getModel("excel").setProperty("/rows", []);
            this.byId("fileUploader").clear();
            this.byId("searchField").setValue("");
            this._file = null;
            this._updateCounts();
            MessageToast.show("Screen cleared.");
        },

    });
});