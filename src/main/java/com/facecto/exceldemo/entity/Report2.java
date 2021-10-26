package com.facecto.exceldemo.entity;

import lombok.Data;
import lombok.experimental.Accessors;

import java.io.Serializable;

/**
 * @author Jon So, https://cto.pub, https://github.com/facecto
 * @version v1.0.0 (2021/10/23)
 */
@Data
@Accessors(chain = true)
public class Report2 implements Serializable {
    private Integer id;
    private String customer;
    private String orderNO;
    private String type;
    private String space;
    private String pay;
    private String date1;
    private String supplier;
    private String bank;
    private String account;
    private String accountNO;
    private String fee1;
    private String time;
}