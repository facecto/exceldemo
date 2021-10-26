package com.facecto.exceldemo.entity;

import lombok.Data;
import lombok.experimental.Accessors;

import java.io.Serializable;
import java.math.BigDecimal;

/**
 * @author Jon So, https://cto.pub, https://github.com/facecto
 * @version v1.0.0 (2021/10/23)
 */
@Data
@Accessors(chain = true)
public class Report1 implements Serializable {
    private Integer id;
    private String customer;
    private String orderNo;
    private String orderType;
    private String space;
    private String pay;
    private String date1;
    private String money;
    private String date2;
    private String date3;
    private BigDecimal money2;
    private BigDecimal money3;
    private String status1;
    private Integer expire;
    private BigDecimal profit;
    private BigDecimal profit1;
    private BigDecimal profit2;
    private BigDecimal profit3;
    private BigDecimal profit4;
    private BigDecimal profit5;
    private String audit;
}
