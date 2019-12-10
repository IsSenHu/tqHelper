package com.tq.helper;

import lombok.Data;

import java.util.Date;

/**
 * @author HuSen
 * create on 2019/12/10
 */
@Data
public class ProjectSuppliesPurchaseOrderItem {
    /** 序号 */
    private Integer number;
    /** 采购单位 */
    private String purchaseUnit;
    /** 网省公司 */
    private String wsgs;
    /** 请购人 */
    private String qgr;
    /** 下单时间 */
    private Date orderTime;
    /** 订单总金额 */
    private Double orderMoney;
    /** 供应商名称 */
    private String gysName;
    /** 付款模式 */
    private String fkms;
    /** 采购类型 */
    private String cglx;
    /** 项目编号 */
    private String xmbh;
    /** 项目名称 */
    private String xmName;
    /** 收货人 */
    private String receiver;
    /** 手机/电话 */
    private String phone;
    /** 收货地址 */
    private String address;
    /** 邮编 */
    private Integer yb;
    /** 备注 */
    private String mark;
}
