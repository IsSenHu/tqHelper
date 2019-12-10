package com.tq.helper;

import lombok.Data;

import java.util.Date;

/**
 * @author HuSen
 * create on 2019/12/10
 */
@Data
class OfflineInvoicingApplicationProjectMaterialPurchaseOrder {
    /** 发票类型 */
    private String fplx;
    /** 发票序号 */
    private Integer fpxh;
    /** 订单号 */
    private Integer orderNum;
    /** 网省公司 */
    private String wsgs;
    /** 购货方名称 */
    private String ghfmc;
    /** 税号 */
    private String sh;
    /** 地址电话 */
    private String addressPhone;
    /** 银行帐号 */
    private String yhzh;
    /** 品类（大类）
     （补录新增字段） */
    private String pl;
    /** 物料名称
     （补录新增字段） */
    private String wlmc;
    /** 商品名称 */
    private String spmc;
    /** 规格型号 */
    private String ggxh;
    /** 单位 */
    private String unit;
    /** 数量 */
    private Double number;
    /** 单价 */
    private Double dj;
    /** 税率 */
    private Double sl;
    /** 金额 */
    private Double money;
    /** 备注 */
    private String mark;
    /** 收票人 */
    private String spr;
    /** 收票人公司名称 */
    private String sprgsmc;
    /** 联系电话 */
    private String lxdh;
    /** 地址 */
    private String address;
    /** 供应商名称 */
    private String gysmc;
    /** 供应商联系人 */
    private String gyslxr;
    /** 供应商联系电话 */
    private String gyslxdh;
    /** 供应商电子邮箱 */
    private String gysdzyj;
    /** 拓展人员邮箱 */
    private String tzryyx;
    /** 供应商折扣率 */
    private Double gyszkl;
    /** 成本价 */
    private Double cbj;
    /** 供应商与英大合同的提供时间 */
    private Date gysydTime;
    /** 进项票寄到英大时间 */
    private Date jxpjdydTime;
}
