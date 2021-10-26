package com.facecto.exceldemo.entity;

import lombok.Data;
import lombok.experimental.Accessors;

import java.math.BigDecimal;

/**
 * @author Jon So, https://cto.pub, https://github.com/facecto
 * @version v1.0.0 (2021/10/23)
 */
@Data
@Accessors(chain = true)
public class Report4 {
    private Integer no;
    private String goods;
    private Integer num;
    private BigDecimal price;
    private BigDecimal tax;
    private String barCode;
    private String pd;
    private String exp;
    private BigDecimal st;
}

