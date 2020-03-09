var i = 0;
var local_uri = 'wss://localhost:8183/';
var count = 1900;
var price = 83770
function doBy() {
    i++;
    $('#txtCount').val(count);
    $('#txtPrice').val(price);
    SaveOrder();
    console.log('doBy');
    if (i < 10)
        doBy();
};
try {
    this.local_websocket = new WebSocket(local_uri);
    this.local_websocket.onopen = function () { alert("open"); console.log('open'); };
    this.local_websocket.onclose = function () { alert("close"); console.log('close'); };
    this.local_websocket.onerror = function () { alert("error"); console.log('error');};
    this.local_websocket.onmessage = function (msg) {
        console.log(msg);
        if (msg.data == 'By') {
            doBy();
        }
        if (msg.data == 'OneBy') {
            $('#txtCount').val(count);
            $('#txtPrice').val(price);
            SaveOrder();
            console.log("OneBy");
        }
    }
} catch (e) {
    console.log(e);
}