var local_uri = 'wss://localhost:8183/';
try {
    this.local_websocket = new WebSocket(local_uri);
    this.local_websocket.onopen = function () { alert("open"); console.log('open'); };
    this.local_websocket.onclose = function () { alert("close"); console.log('close'); };
    this.local_websocket.onerror = function () { alert("error"); console.log('error'); };
    this.local_websocket.onmessage = function (msg) {
        console.log(msg);
        var sellBuyInfo = msg.data.split(',');
        if (sellBuyInfo[0] == 'OneBuy') {
            $('input:radio:checked').val('65');
        }
        if (sellBuyInfo[0] == 'OneSell') {
            $('input:radio:checked').val('86');
        }
        var marketName = sellBuyInfo[1];
        var count = sellBuyInfo[2];
        var price = sellBuyInfo[3];
        $('#txtCount').val(count);
        $('#txtPrice').val(price);
        $.get("https://online.mobinsb.com/StockAutoCompleteHandler.ashx?ShowAll=0&MarketType=0&lan=fa&ShowNotApproved=0&q=" + marketName + "&.rand=7a1ee3eb26d046e8bc5c97dba4830217", function (data) {
            var arr = data.split('\n');
            var a = new Object();
            a.innervalue = arr[0];
            var temp = a.innervalue.split(',');
            a.display = temp[0] + ' --- ' + temp[2];
            a.value = temp[2];
            $('#hiddrpExchangeList').val(a.innervalue);
            callbackF(a);
            console.log(a);
            setTimeout(function () {
                $('#txtCount').val(count);
                $('#txtPrice').val(price);
                SaveOrder();
                console.log("Done");
            }, 500);
        });
    }
} catch (e) {
    console.log(e);
}