module.exports = function(RED){
    function convertxlsx(config){
        RED.nodes.createNode(this,config);
        //Los inputs
        this.filepath = config.filepath;
        this.rangecell = config.rangecell;
        this.columnkey = config.columnkey;
        this.sheet = config.sheet;

        var node = this;
        node.on('input', function(msg){
            const filepath = msg.filepath || node.filepath;
            const rangecell = msg.rangecell || node.rangecell;
            const columnkey = msg.columnkey || node.columnkey;
            const sheet = msg.sheet || node.sheet;
        
            const excelToJson = require('convert-excel-to-json');
            let result = {};
            let columnas;
            const defaultColumn = '{"*": "{{columnHeader}}"}';
        
            if (columnkey) {
                columnas = JSON.parse(columnkey);
            } else {
                columnas = JSON.parse(defaultColumn);
            }
        
            result = excelToJson({
                sourceFile: filepath,
                range: rangecell,
                columnToKey: columnas
            });
        
            if (sheet) {
                msg.payload = result[sheet];
            } else {
                msg.payload = result;
            }
        
            node.send(msg);
        });
        
    }
    RED.nodes.registerType("XLSX-to-json", convertxlsx);
}
