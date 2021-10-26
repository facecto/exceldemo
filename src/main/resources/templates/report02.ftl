<table style="border-style:thin;text-align: center;vertical-align: middle;">
    <caption>${sheetName}</caption>
    <thead>
    <tr style="background-color:#f2f2f2">
        <th colspan="13" style="text-align: center;vertical-align: middle;font-weight: bold;font-size: 14px;">${title}</th>
    </tr>
    <tr style="background-color:#f2f2f2">
        <th colspan="9" style="text-align: right">${tableMan}</th>
        <th colspan="4" style="text-align: right">${tableTime}</th>
    </tr>
    <tr>
        <#list titles as title>
            <th>${title}</th>
        </#list>
    </tr>
    </thead>
    <tbody>
    <#list data as item>
        <tr>
            <td>${item.id}</td>
            <td>${item.customer}</td>
            <td>${item.orderNO}</td>
            <td>${item.type}</td>
            <td>${item.space}</td>

            <td>${item.pay}</td>
            <td>${item.date1}</td>

            <td>${item.supplier}</td>
            <td>${item.bank}</td>
            <td>${item.account}</td>
            <td>${item.accountNO}</td>
            <td>${item.fee1}</td>
            <td>${item.time}</td>
        </tr>
    </#list>
    </tbody>
</table>