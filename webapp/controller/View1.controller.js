sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/m/MessageToast",
    "sap/m/MessageBox",
    "sap/ui/model/json/JSONModel",
    "sap/ui/core/BusyIndicator",
    "sap/ui/model/Filter",
    "sap/ui/model/FilterOperator"
], function (
    Controller,
    MessageToast,
    MessageBox,
    JSONModel,
    BusyIndicator,
    Filter,
    FilterOperator
) {
    "use strict";

    return Controller.extend("me11uploadingdata.controller.View1", {

        
        //CONSTANTS                                        
        
        CONST: {
            GROUP_ID: "upload",
            ENTITY_SET: "/UploadHeaderSet",
            MAX_ROWS: 1000,
            STATUS_READY: "Ready",
            STATUS_INVALID: "Invalid",
            STATUS_SUCCESS: "Success",
            STATUS_ERROR: "Error"
        },

        onInit: function () {
            this._file = null;

            this.getView().setModel(new JSONModel({
                rows: []
            }), "excel");

            this._configureODataModel();
            this._loadXLSXLibrary();
        },

        _configureODataModel: function () {
            var oModel = this.getOwnerComponent().getModel();

            if (oModel) {
                oModel.setUseBatch(true);
                oModel.setDeferredGroups([this.CONST.GROUP_ID]);
            }
        },

      
        // LOGIN USER    
        
        _getLoggedInUser: function () {
            try {
                return sap.ushell.Container.getUser()
                .getEmail()
                .split("@")[0].toUpperCase();
            } catch (e) {
                return "";
            }
        },
         // XLSX LIB                                        
        
        _loadXLSXLibrary: function () {
            if (typeof XLSX !== "undefined") {
                return;
            }

            var script = document.createElement("script");
            script.src = sap.ui.require.toUrl(
                "me11uploadingdata/libs/xlsx.full.min.js"
            );

            script.onerror = function () {
                MessageBox.error("Unable to load Excel library.");
            };

            document.head.appendChild(script);
        },

        
        // FILE SELECT                                    
        
        onFileChange: function (oEvent) {
            var aFiles = oEvent.getParameter("files");

            if (aFiles && aFiles.length > 0) {
                this._file = aFiles[0];
                MessageToast.show("File Selected: " + this._file.name);
            }
        },

        
        //READ EXCEL                                        
       
        onReadExcel: function () {
            var that = this;

            if (!this._file) {
                MessageBox.warning("Please select Excel file.");
                return;
            }

            if (typeof XLSX === "undefined") {
                MessageBox.error("Excel library not loaded.");
                return;
            }

            BusyIndicator.show(0);

            var reader = new FileReader();

            reader.onload = function (oEvent) {
                try {
                    var workbook = XLSX.read(
                        new Uint8Array(oEvent.target.result),
                        { type: "array" }
                    );

                    var sheet =
                        workbook.Sheets[workbook.SheetNames[0]];

                    var rawData = XLSX.utils.sheet_to_json(sheet, {
                        header: 1,
                        defval: ""
                    });

                    that._processExcelRows(rawData);

                } catch (e) {
                    MessageBox.error("Excel read error.");
                } finally {
                    BusyIndicator.hide();
                }
            };

            reader.readAsArrayBuffer(this._file);
        },

        _processExcelRows: function (rawData) {
            var that = this;
            var aRows = [];
            var oDuplicateMap = {};

            rawData.slice(1).forEach(function (row, index) {

                var oRow = that._mapExcelRow(row, index);

                if (that._isBlankRow(oRow)) {
                    return;
                }

                var aErrors = that._validateRow(oRow);

                var sKey = [
                    oRow.Lifnr,
                    oRow.Matnr,
                    oRow.Ekorg,
                    oRow.Werks,
                    oRow.Datab,
                    oRow.Datbi
                ].join("|");

                if (oDuplicateMap[sKey]) {
                    aErrors.push("Duplicate row in Excel");
                }

                oDuplicateMap[sKey] = true;

                that._setRowStatus(oRow, aErrors);

                aRows.push(oRow);
            });

            if (aRows.length > this.CONST.MAX_ROWS) {
                this._clearRows();
                MessageBox.warning(
                    "Maximum " + this.CONST.MAX_ROWS +
                    " rows allowed."
                );
                return;
            }

            this._setRows(aRows);
            MessageToast.show("Excel loaded successfully.");
        },

        _mapExcelRow: function (row, index) {
            return {
                RowNo: index + 1,
                Lifnr: this._toString(row[0]),
                Matnr: this._toString(row[1]),
                Ekorg: this._toString(row[2]),
                Werks: this._toString(row[3]),
                Normb: this._toString(row[4]),
                Meins: this._toString(row[5]),
                Aplfz: this._toString(row[6]),
                Norbm: this._toDecimal(row[7]),
                Mwskz: this._toString(row[8]),
                Netpr: this._toDecimal(row[9]),
                Peinh: this._toInteger(row[10]),
                Bprme: this._toString(row[11]),
                Datab: this._toDate(row[12]),
                Datbi: this._toDate(row[13]),
                Kschl_02: this._toString(row[14]),
                Kbetr_02: this._toString(row[15]),
                Kschl_03: this._toString(row[16]),
                Kbetr_03: this._toString(row[17]),
                StatusText: "",
                StatusState: "",
                ValidationMessage: ""
            };
        },

       
        // SEARCH                                            
      
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

            oBinding.filter(new Filter({
                filters: [
                    new Filter("Lifnr", FilterOperator.Contains, sValue),
                    new Filter("Matnr", FilterOperator.Contains, sValue),
                    new Filter("Werks", FilterOperator.Contains, sValue),
                    new Filter("StatusText", FilterOperator.Contains, sValue),
                    new Filter("ValidationMessage", FilterOperator.Contains, sValue)
                ],
                and: false
            }));
        },

        
        // SEND BACKEND                                      
        
        onSendToBackend: function () {
            this._revalidateAllRows();

            var aValidRows = this._getRows().filter(function (row) {
                return row.StatusText === "Ready";
            });

            if (aValidRows.length === 0) {
                MessageBox.warning("No valid rows to upload.");
                return;
            }

            this._sendUpload(aValidRows);
        },

        _sendUpload: function (aRows) {
            var that = this;
            var oModel = this.getView().getModel();
            var sUser = this._getLoggedInUser();
           

            if (!sUser) {
                MessageBox.error("Unable to fetch login user.");
                return;
            }

            BusyIndicator.show(0);

            oModel.create(
                this.CONST.ENTITY_SET,
                this._buildPayload(aRows, sUser),
                {
                    groupId: this.CONST.GROUP_ID,

                    success: function (oData) {
                        BusyIndicator.hide();
                        that._applyBackendResult(oData);
                        that.byId("fileUploader").clear();
                        that._file = null;

                        MessageBox.success(
                            "Upload completed successfully."
                        );
                    },

                    error: function (oErr) {
                        BusyIndicator.hide();
                        MessageBox.error(
                            "Upload failed. Please contact support."
                        );
                    }
                }
            );

            oModel.submitChanges({
                groupId: this.CONST.GROUP_ID
            });
        },

        _buildPayload: function (aRows, sUser) {
          
            var that = this;
           

            return {
                UploadId: sUser,
                ToItems: aRows.map(function (row) {
                    return {
                        UploadId: sUser,
                        Lifnr: row.Lifnr,
                        Matnr: row.Matnr,
                        Ekorg: row.Ekorg,
                        Werks: row.Werks,
                        Normb: row.Normb,
                        Meins: row.Meins,
                        Aplfz: row.Aplfz,
                        Norbm: that._formatDecimal(row.Norbm, 3),
                        Mwskz: row.Mwskz,
                        Netpr: that._formatDecimal(row.Netpr, 2),
                        Peinh: String(row.Peinh),
                        Bprme: row.Bprme,
                        Datab: row.Datab,
                        Datbi: row.Datbi,
                        Kschl_02: row.Kschl_02,
                        Kbetr_02: row.Kbetr_02,
                        Kschl_03: row.Kschl_03,
                        Kbetr_03: row.Kbetr_03
                    };
                })
            };
        },

        _applyBackendResult: function (oData) {
            if (!oData ||
                !oData.ToItems ||
                !oData.ToItems.results) {
                return;
            }

            var aBackend = oData.ToItems.results;
            var aRows = this._getRows();

            aRows.forEach(function (row) {

                var oMatch = aBackend.find(function (item) {
                    return item.Lifnr === row.Lifnr &&
                           item.Matnr === row.Matnr;
                });

                if (oMatch) {
                    row.StatusText =
                        oMatch.Status === "S"
                            ? "Success"
                            : "Error";

                    row.StatusState =
                        oMatch.Status === "S"
                            ? "Success"
                            : "Error";

                    row.ValidationMessage = oMatch.Message;
                }
            });

            this.getView().getModel("excel").refresh(true);
            this._updateCounts();
        },

        
        // VALIDATION                                     
        
        _validateRow: function (row) {
            var aErrors = [];

            if (!row.Lifnr) { aErrors.push("LIFNR missing"); }
            if (!row.Matnr) { aErrors.push("MATNR missing"); }
            if (!row.Ekorg) { aErrors.push("EKORG missing"); }
            if (!row.Werks) { aErrors.push("WERKS missing"); }
            if (!row.Meins) { aErrors.push("MEINS missing"); }
            if (!row.Mwskz) { aErrors.push("MWSKZ missing"); }

            if (!row.Norbm || Number(row.Norbm) <= 0) {
                aErrors.push("NORBM > 0 required");
            }

            if (!row.Peinh || Number(row.Peinh) <= 0) {
                aErrors.push("PEINH > 0 required");
            }

            if (!/^\d{8}$/.test(row.Datab)) {
                aErrors.push("DATAB invalid");
            }

            if (!/^\d{8}$/.test(row.Datbi)) {
                aErrors.push("DATBI invalid");
            }

            if (row.Datab > row.Datbi) {
                aErrors.push("DATAB > DATBI");
            }

            return aErrors;
        },

        _revalidateAllRows: function () {
            var that = this;

            this._getRows().forEach(function (row) {
                that._setRowStatus(
                    row,
                    that._validateRow(row)
                );
            });

            this.getView().getModel("excel").refresh(true);
        },

        _setRowStatus: function (row, aErrors) {
            if (aErrors.length > 0) {
                row.StatusText = this.CONST.STATUS_INVALID;
                row.StatusState = this.CONST.STATUS_ERROR;
                row.ValidationMessage = aErrors.join(", ");
            } else {
                row.StatusText = this.CONST.STATUS_READY;
                row.StatusState = this.CONST.STATUS_SUCCESS;
                row.ValidationMessage = "OK";
            }
        },

       
        // DOWNLOAD                                           
      
        onDownloadExcel: function () {
            var aRows = this._getRows();

            if (aRows.length === 0) {
                MessageBox.warning("No data available.");
                return;
            }

            var aData = [["Supplier", "Material", "Status", "Message"]];

            aRows.forEach(function (row) {
                aData.push([
                    row.Lifnr,
                    row.Matnr,
                    row.StatusText,
                    row.ValidationMessage
                ]);
            });

            var wb = XLSX.utils.book_new();
            var ws = XLSX.utils.aoa_to_sheet(aData);

            XLSX.utils.book_append_sheet(wb, ws, "Result");
            XLSX.writeFile(wb, "ME11_Upload_Result.xlsx");
        },

        
        // CLEAR                                              
        
        onClear: function () {
            this._clearRows();

            this.byId("fileUploader").clear();
            this.byId("searchField").setValue("");

            this._file = null;

            MessageToast.show("Screen cleared.");
        },

        _clearRows: function () {
            this._setRows([]);
        },

        
        // MODEL HELPERS                                       
        
        _getRows: function () {
            return this.getView()
                .getModel("excel")
                .getProperty("/rows");
        },

        _setRows: function (aRows) {
            this.getView()
                .getModel("excel")
                .setProperty("/rows", aRows);

            this._updateCounts();
        },

        _updateCounts: function () {
            var aRows = this._getRows();

            this.byId("txtRows")
                .setText("Rows: " + aRows.length);

            this.byId("txtValid")
                .setText(
                    "Valid: " +
                    aRows.filter(function (row) {
                        return row.StatusText === "Ready";
                    }).length
                );

            this.byId("txtInvalid")
                .setText(
                    "Invalid: " +
                    aRows.filter(function (row) {
                        return row.StatusText === "Invalid";
                    }).length
                );
        },

        
        // CONVERTERS                                          
        
        _toString: function (v) {
            return v ? String(v).trim() : "";
        },

        _toDecimal: function (v) {
            return v ? String(v).replace(/,/g, "") : "";
        },

        _toInteger: function (v) {
            return v ? parseInt(v, 10) : "";
        },

        _formatDecimal: function (v, scale) {
            return parseFloat(v || 0).toFixed(scale);
        },

        _isBlankRow: function (row) {
            return !(
                row.Lifnr ||
                row.Matnr ||
                row.Ekorg ||
                row.Werks
            );
        },

       _toDate: function (v) {
            if (!v) {
                return "";
            }

            if (typeof v === "number") {
                var d1 = new Date(Math.round((v - 25569) * 86400 * 1000));
                return this._formatDate(d1);
            }

            var s = String(v).trim();

            if (/^\d{8}$/.test(s)) {
                return s;
            }

            var sep = "";

            if (s.indexOf(".") > -1) {
                sep = ".";
            } else if (s.indexOf("-") > -1) {
                sep = "-";
            } else if (s.indexOf("/") > -1) {
                sep = "/";
            }

            if (sep) {
                var p = s.split(sep);

                if (p.length === 3) {
                    var d2 = new Date(p[2], p[1] - 1, p[0]);

                    if (!isNaN(d2.getTime())) {
                        return this._formatDate(d2);
                    }
                }
            }

            var d3 = new Date(s);

            if (!isNaN(d3.getTime())) {
                return this._formatDate(d3);
            }

            return "";
        },
        _formatDate: function (d) {
            return d.getFullYear() +
                String(d.getMonth() + 1).padStart(2, "0") +
                String(d.getDate()).padStart(2, "0");
        }
    });
});