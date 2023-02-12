module.exports = function (App) {
  const { KLNode, KLError } = require('@kumologica/devkit');
  const XLSXWrite = require("json-as-xlsx");
  const XLSXRead = require("xlsx");
  class ExcelError extends KLError { }

  class Excel extends KLNode {
    constructor(props) {
      super(App, props);
      this.operation = props.operation;
      this.content = props.content;

      // Method bindings
      this.handle = this.handle.bind(this);
    }

    async handle(msg) {
      try {
        let Content = App.util.evaluateDynamicField(this.content, msg, this);
        if (this.operation === 'read') {
          let wb =  XLSXRead.read(Content);
          let sheet_name_list = wb.SheetNames;
          let xlData = [];
          sheet_name_list.forEach(element => {
            let sheet = {
              "sheet": element,
              content:  XLSXRead.utils.sheet_to_json(wb.Sheets[element])
            }
            xlData.push(sheet)
          });
          msg.payload = xlData
          this.send(msg)
          return
        } else {
          let settings = {
            extraLength: 3,
            writeMode: "write",
            writeOptions: {"type" : "buffer"},
            RTL: false, 
          }
          msg.payload  = XLSXWrite(Content, settings)
          this.send(msg)
          return
        }
      } catch (error) {
        this.sendError(new ExcelError(error), msg)
        return;
      }
    }
  }
  App.nodes.registerType('Excel', Excel);
};
