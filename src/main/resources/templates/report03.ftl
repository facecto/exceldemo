<table style="border-style:thin;text-align: center;vertical-align: middle;">
    <caption>${sheetName}</caption>
    <thead>
    <tr style="background-color:#e2e2e2">
        <th colspan="25" style="text-align: center;vertical-align: middle;font-weight: bold;font-size: 14px;">${title}</th>
    </tr>
    <tr style="background-color:#e2e2e2">
        <th colspan="21" style="text-align: right">${tableMan}</th>
        <th colspan="4" style="text-align: right">${tableTime}</th>
    </tr>
    <tr style="background-color:#e2e2e2">
        <th colspan="7">OrderInfo</th>
        <th colspan="18">Profit</th>
    </tr>
    <tr style="background-color:#e2e2e2">
        <th rowspan="4">No.</th>
        <th rowspan="4">OrderNo</th>
        <th rowspan="4">Customer</th>
        <th rowspan="4">Date</th>
        <th rowspan="4">OrderType</th>
        <th rowspan="4">OrderMoney</th>
        <th rowspan="4">Profit</th>
        <th colspan="3">PartA</th>
        <th colspan="4">PartB</th>
        <th colspan="11">PartC</th>
    </tr>
    <tr style="background-color:#e2e2e2">
        <th rowspan="2">One 10%</th>
        <th colspan="2" rowspan="2">Two 30%</th>
        <th colspan="2" rowspan="2">Three 20%</th>
        <th colspan="2" rowspan="2">Four 40%</th>
        <th rowspan="2">Five 15%</th>
        <th colspan="9">Six 65%</th>
        <th rowspan="2">Seven 20%</th>
    </tr>
    <tr style="background-color:#e2e2e2">
        <th colspan="2">Eight 45%</th>
        <th colspan="2">Night 5%</th>
        <th colspan="2">Ten 10%</th>
        <th colspan="2">Eleven 25%</th>
        <th>Twelve 15%</th>
    </tr>
    <tr style="background-color:#e2e2e2">
        <th>Num</th>
        <th>Name</th>
        <th>Num</th>
        <th>Name</th>
        <th>Num</th>
        <th>Name</th>
        <th>Num</th>
        <th>Num</th>
        <th>Name</th>
        <th>Num</th>
        <th>Name</th>
        <th>Num</th>
        <th>Name</th>
        <th>Num</th>
        <th>Name</th>
        <th>Num</th>
        <th>Num</th>
        <th>Num</th>
    </tr>
    </thead>
    <tbody>
    <#list data as item>
        <tr>
            <td>${item.id}</td>
            <td>${item.orderNo}</td>
            <td>${item.customer}</td>
            <td>${item.date1}</td>
            <td>${item.orderType}</td>

            <td>${item.orderMoney}</td>
            <td>${item.profit0}</td>
            <td>${item.profit1}</td>
            <td>${item.name1}</td>

            <td>${item.profit2}</td>
            <td>${item.name2}</td>
            <td>${item.profit3}</td>
            <td>${item.name3}</td>
            <td>${item.profit4}</td>

            <td>${item.profit5}</td>
            <td>${item.name4}</td>
            <td>${item.profit6}</td>
            <td>${item.name5}</td>
            <td>${item.profit7}</td>

            <td>waiter</td>
            <td>${item.profit8}</td>
            <td>${item.name6}</td>
            <td>${item.profit9}</td>
            <td>${item.profit10}</td>

            <td>${item.profit11}</td>
        </tr>
    </#list>
    </tbody>
</table>