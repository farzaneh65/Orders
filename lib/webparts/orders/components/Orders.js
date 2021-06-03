var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import * as React from 'react';
import styles from './Orders.module.scss';
import { SPHttpClient } from '@microsoft/sp-http';
import { Web } from "sp-pnp-js/lib/pnp";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { FilePicker } from '@pnp/spfx-controls-react/lib/FilePicker';
import { TextField, autobind, PrimaryButton, Dropdown } from 'office-ui-fabric-react';
var textFieldStyles = { fieldGroup: { width: 500 } };
var narrowTextFieldStyles = { fieldGroup: { width: 100 } };
var narrowDropdownStyles = { dropdown: { width: 300 } };
var Orders = /** @class */ (function (_super) {
    __extends(Orders, _super);
    function Orders(props, state) {
        var _this = _super.call(this, props) || this;
        ////////// Component Did Mount ////////////////
        _this.options = [];
        _this.state = {
            status: "Ready",
            orderListItem: {
                Id: 0,
                Title: "",
                Quantity: 0,
                DestinationId: 1,
                Description: "",
                LinkToFile: "",
                Owners: []
            },
            filePickerResult: null,
            Users: [],
            message: ""
        };
        return _this;
    }
    ////////// Get Destinations ////////////////
    Orders.prototype._getDestinations = function () {
        var url = this.props.siteURL + "/_api/web/lists/getbytitle('Destinations')/Items?$filter=Active eq 1";
        return this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
            .then(function (response) {
            return response.json();
        })
            .then(function (json) {
            return json.value;
        });
    };
    ////////// Get Orders-For debugging ////////////////
    Orders.prototype._getOrderss = function () {
        var url = this.props.siteURL + "/_api/web/lists/getbytitle('Orders')/Items";
        return this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
            .then(function (response) {
            return response.json();
        }).then(function (json) {
            console.log(json.value);
            return json.value;
        });
    };
    ////////// Upload Recieved File To Document Library ////////////////  
    Orders.prototype.uploadFileToLibrary = function (ItemId) {
        return __awaiter(this, void 0, void 0, function () {
            var web;
            var _this = this;
            return __generator(this, function (_a) {
                web = new Web(this.props.siteURL);
                web.folders.getByName('OrdersFiles').folders.add("Order-" + ItemId)
                    .then(function (fl) { return __awaiter(_this, void 0, void 0, function () {
                    var _a, _b, _c;
                    return __generator(this, function (_d) {
                        switch (_d.label) {
                            case 0:
                                _b = (_a = web.getFolderByServerRelativeUrl("OrdersFiles/Order-" + ItemId).files).add;
                                _c = [this.state.filePickerResult[0].fileName];
                                return [4 /*yield*/, this.state.filePickerResult[0].downloadFileContent()];
                            case 1:
                                _b.apply(_a, _c.concat([_d.sent(), true]))
                                    .then(function (result) {
                                    console.log(result);
                                    return true;
                                });
                                return [2 /*return*/];
                        }
                    });
                }); })
                    .catch(function (error) {
                    console.log(error.errorMessage);
                    return false;
                });
                return [2 /*return*/];
            });
        });
    };
    Orders.prototype.componentDidMount = function () {
        return __awaiter(this, void 0, void 0, function () {
            var destinations;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this._getDestinations()];
                    case 1:
                        destinations = _a.sent();
                        console.log(this);
                        destinations.forEach(function (dest) {
                            _this.options.push({
                                key: dest.ID,
                                text: dest.Title
                            });
                        });
                        return [2 /*return*/];
                }
            });
        });
    };
    ////////// Add Item To the Order List ////////////////  
    Orders.prototype.btnAdd_click = function () {
        var _this = this;
        this.state.Users.forEach(function (element) {
            console.log(element);
        });
        var web = new Web(this.props.siteURL);
        web.lists.getById('B1C37CDB-2E4D-41F5-8381-B22701ABA693').items
            .add({
            Title: this.state.orderListItem.Title,
            Quantity: this.state.orderListItem.Quantity,
            DestinationId: this.state.orderListItem.DestinationId,
            Description: this.state.orderListItem.Description,
            OwnersId: {
                "results": this.state.Users
            }
        })
            .then(function (r) { return __awaiter(_this, void 0, void 0, function () {
            var _a, _b, _c, fileUrl;
            return __generator(this, function (_d) {
                switch (_d.label) {
                    case 0:
                        console.log(r);
                        _b = (_a = r.item.attachmentFiles).add;
                        _c = [this.state.filePickerResult[0].fileName];
                        return [4 /*yield*/, this.state.filePickerResult[0].downloadFileContent()];
                    case 1:
                        _b.apply(_a, _c.concat([_d.sent()]));
                        console.log(r.data.Id);
                        if (this.uploadFileToLibrary(r.data.Id)) {
                            fileUrl = this.props.siteURL + ("/OrdersFiles/Order-" + r.data.Id + "/" + this.state.filePickerResult[0].fileName);
                            console.log(fileUrl);
                            r.item.update({
                                LinkToFile: {
                                    "__metadata": { type: "SP.FieldUrlValue" },
                                    Description: this.state.filePickerResult[0].fileNameWithoutExtension,
                                    Url: fileUrl
                                }
                            })
                                .then((function (i) {
                                alert('Order is registered successfully.');
                            }))
                                .catch(function (err) {
                                console.log(err.errorMessage);
                            });
                        }
                        return [2 /*return*/];
                }
            });
        }); })
            .catch(function (error) {
            console.log(error.errorMessage);
        });
    };
    Orders.prototype._getPeoplePickerItems = function (items) {
        ///console.log('Items:', items);
        var getSelectedUsers = [];
        for (var item in items) {
            getSelectedUsers.push(items[item].id);
        }
        this.setState({ Users: getSelectedUsers });
    };
    Orders.prototype.render = function () {
        var _this = this;
        var dropdownRef = React.createRef();
        return (React.createElement("div", { className: styles.orders },
            React.createElement(TextField, { label: 'Item', required: true, value: (this.state.orderListItem.Title).toString(), styles: textFieldStyles, onChanged: function (e) { _this.state.orderListItem.Title = e; } }),
            React.createElement(TextField, { label: 'Quantity', required: true, type: 'number', 
                //value={(this.state.orderListItem.Quantity)}
                styles: textFieldStyles, onChanged: function (e) { _this.state.orderListItem.Quantity = e; } }),
            React.createElement(Dropdown, { componentRef: dropdownRef, placeholder: "Select an option", label: 'Destination', options: this.options, defaultSelectedKey: "", required: true, styles: narrowDropdownStyles, onChanged: function (e) { _this.state.orderListItem.DestinationId = Number(e.key); } }),
            React.createElement(PeoplePicker, { context: this.props.context, titleText: "Owners", personSelectionLimit: 3, showtooltip: true, required: true, disabled: false, 
                //selectedItems={this._getPeoplePickerItems} 
                onChange: this._getPeoplePickerItems, showHiddenInUI: false, ensureUser: true, principalTypes: [PrincipalType.User], resolveDelay: 1000 }),
            React.createElement(TextField, { label: 'Description', type: 'string', 
                //value={(this.state.orderListItem.Quantity)}
                styles: textFieldStyles, onChanged: function (e) { _this.state.orderListItem.Quantity = e; } }),
            React.createElement("p", null),
            React.createElement(FilePicker, { buttonLabel: "Select File", onSave: function (filePickerResult) {
                    _this.setState({ filePickerResult: filePickerResult }); //alert('File has been uploaded successfully.');
                    document.getElementById('msg').innerHTML = filePickerResult[0].fileName + ' has been uploaded successfully!!';
                    console.log(_this.state.filePickerResult);
                }, onChange: function (filePickerResult) {
                    _this.setState({ filePickerResult: filePickerResult }); //alert('File has been uploaded successfully!!');
                    document.getElementById('msg').innerHTML = filePickerResult[0].fileName + ' has been uploaded successfully!!';
                    console.log(_this.state.filePickerResult);
                }, context: this.props.context }),
            React.createElement("p", { id: "msg", style: { color: 'green' } }),
            React.createElement("p", { className: styles.title },
                React.createElement(PrimaryButton, { style: { backgroundColor: 'rgb(34, 122, 138)' }, text: 'Add', title: 'Add', onClick: this.btnAdd_click }))));
    };
    __decorate([
        autobind
    ], Orders.prototype, "btnAdd_click", null);
    __decorate([
        autobind
    ], Orders.prototype, "_getPeoplePickerItems", null);
    return Orders;
}(React.Component));
export default Orders;
//# sourceMappingURL=Orders.js.map