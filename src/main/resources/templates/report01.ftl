<table style="border-style:thin;text-align: center;vertical-align: middle;">
    <caption>${sheetName}</caption>
    <thead>
    <tr style="background-color:#f2f2f2">
        <th colspan="21" style="font-weight: bold;font-size: 14px;">${title}</th>
    </tr>
    <tr style="background-color:#f2f2f2">
        <th colspan="17" style="text-align: right">${tableMan}</th>
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
            <td>${item.orderNo}</td>
            <td>${item.orderType}</td>
            <td>${item.space}</td>
            <td>${item.pay}</td>
            <td>${item.date1}</td>
            <td>${item.money}</td>
            <td>${item.date2}</td>
            <td>${item.date3}</td>
            <td>${item.money2}</td>
            <td>${item.money3}</td>
            <td>${item.status1}</td>
            <td>${item.expire}</td>
            <td>${item.profit}</td>
            <td>${item.profit1}</td>
            <td>${item.profit2}</td>
            <td>${item.profit3}</td>
            <td>${item.profit4}</td>
            <td>${item.profit5}</td>
            <td>${item.audit}</td>
        </tr>
    </#list>
    </tbody>
</table>