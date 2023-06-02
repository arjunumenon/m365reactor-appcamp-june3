const axios = require("axios");
const querystring = require("querystring");
const { TeamsActivityHandler, CardFactory } = require("botbuilder");
const { SupplierME } = require('./messaginExtensions/supplierME.js')

class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();
  }

  // Message extension Code
  // Search.
  async handleTeamsMessagingExtensionQuery(context, query) {

    const queryName = query.parameters[0].name;
    const searchQuery = query.parameters[0].value;

    switch (queryName) {
      case "searchQuery":  // Search for suppliers
        return await SupplierME.handleTeamsMessagingExtensionQuery(context, searchQuery);
      default:
        return null;
    }
  }

  async handleTeamsMessagingExtensionSelectItem(context, item) {

    switch (item.queryType) {
      case "supplierME":  // Search for suppliers
        return SupplierME.handleTeamsMessagingExtensionSelectItem(context, item);
      default:
        return null;
    }

  }
}

module.exports.TeamsBot = TeamsBot;
