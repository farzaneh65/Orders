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
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'OrdersWebPartStrings';
import Orders from './components/Orders';
var OrdersWebPart = /** @class */ (function (_super) {
    __extends(OrdersWebPart, _super);
    function OrdersWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    OrdersWebPart.prototype.render = function () {
        var element = React.createElement(Orders, {
            description: this.properties.description,
            siteURL: this.context.pageContext.web.absoluteUrl,
            context: this.context
        });
        ReactDom.render(element, this.domElement);
    };
    OrdersWebPart.prototype.onDispose = function () {
        ReactDom.unmountComponentAtNode(this.domElement);
    };
    Object.defineProperty(OrdersWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: true,
        configurable: true
    });
    OrdersWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneTextField('description', {
                                    label: strings.DescriptionFieldLabel
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return OrdersWebPart;
}(BaseClientSideWebPart));
export default OrdersWebPart;
//# sourceMappingURL=OrdersWebPart.js.map