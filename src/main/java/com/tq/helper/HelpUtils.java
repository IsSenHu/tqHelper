package com.tq.helper;

import com.alibaba.fastjson.JSON;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.InputStream;
import java.text.DecimalFormat;
import java.util.*;
import java.util.function.Predicate;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * @author HuSen
 * create on 2019/12/10
 */
@Slf4j
public class HelpUtils {
    private static final String N1 = "订单序号必填";
    private static final String N2 = "采购单位必填";
    private static final String N3 = "网省公司必填";
    private static final String N4 = "请购人必填";
    private static final String N5 = "下单时间必填";
    private static final String N6 = "订单总额必填，1亿以内，至多两位小数";
    private static final String N7 = "供应商名称必填";
    private static final String N8 = "付款模式必填，支持先货后款和先款后货";
    private static final String N9 = "采购类型必填，支持其他物资类、车辆类、计算机类、集采物资类、专业物资类";
    private static final String N10 = "项目编号非必填，仅支持导入一条";
    private static final String N11 = "项目名称非必填，项目名称字符长度不能超过118";
    private static final String N12 = "收货人必填，长度不超过10，仅支持汉字、数字";
    private static final String N13 = "手机/电话必填，只能是11位";
    private static final String N14 = "收货地址必填，长度检验不超过200";
    private static final String N15 = "邮编必填，长度检验6位，仅支持数字";
    private static final String N16 = "备注非必填，长度检验不超过200";

    private static final String M1 = "发票序号必填，对应的供应商得和分页一订单序号对应的供应商一致";
    private static final String M2 = "订单号必填，按照项目物资采购订单表订单序号填列";
    private static final String M3 = "网省公司必填";
    private static final String M4 = "购货方名称必填";
    private static final String M5 = "税号必填，要18位";
    private static final String M6 = "地址电话必填，要同时有地址和电话";
    private static final String M7 = "银行账号必填，要同时又开户行和帐号";
    private static final String M8 = "品类必填，大类名称，不得超过20个汉字";
    private static final String M9 = "物料名称必填，若集体企业无特殊要求，与小类名称或商品名称填写一致即可，不得超过20个汉字";
    private static final String M10 = "商品名称必填，货物或应税劳务、服务名称，不能超过45个汉字";
    private static final String M11 = "规格型号必填，如需此信息请务必填写，不得超过20个汉字";
    private static final String M12 = "单位必填";
    private static final String M13 = "数量必填";
    private static final String M14 = "单价必填";
    private static final String M15 = "税率必填";
    private static final String M16 = "金额必填";
    private static final String M17 = "收票人必填";
    private static final String M18 = "收票人公司名称必填";
    private static final String M19 = "联系电话必填";
    private static final String M20 = "地址必填";
    private static final String M21 = "供应商名称必填";
    private static final String M22 = "供应商联系人必填";
    private static final String M23 = "供应商联系电话必填";
    private static final String M24 = "供应商电子邮箱必填";
    private static final String M25 = "拓展人员邮箱必填";
    private static final String M26 = "供应商折扣率必填";
    private static final String M27 = "成本价必填";
    private static final String M28 = "供应商与英大合同的提供时间必填";
    private static final String M29 = "进项票寄到英大时间必填";
    private static final String M30 = "供应商联系人、供应商联系电话、供应商电子邮箱一一对应";
    private static final String M31 = "F列和U列公司相同";

    private static final Set<String> PAY_MODES = new HashSet<>(2);
    private static final Set<String> PURCHASE_MODES = new HashSet<>(5);

    static {
        PAY_MODES.add("先货后款");
        PAY_MODES.add("先款后货");

        PURCHASE_MODES.add("物资类");
        PURCHASE_MODES.add("车辆类");
        PURCHASE_MODES.add("计算机类");
        PURCHASE_MODES.add("集采物资类");
        PURCHASE_MODES.add("专业物资类");
    }

    static JsonResult<Valid> doHelp(InputStream in) {
        try {
            XSSFWorkbook workbook = new XSSFWorkbook(in);
            // 项目物资采购订单模版
            XSSFSheet projectSuppliesPurchaseOrderTemplate = workbook.getSheetAt(0);
            // 项目物资采购订单线下开票申请表
            XSSFSheet offlineInvoicingApplicationProjectMaterialPurchaseOrder = workbook.getSheetAt(1);

            XSSFRow rowHead = projectSuppliesPurchaseOrderTemplate.getRow(1);
            int physicalNumberOfCells = rowHead.getPhysicalNumberOfCells();
            if (physicalNumberOfCells != 17) {
                return new JsonResult<>(1, "文件不遵循模版规范", null);
            }
            JsonResult<List<ProjectSuppliesPurchaseOrderItem>> purchaseOrderItems = projectSuppliesPurchaseOrderItems(projectSuppliesPurchaseOrderTemplate);
            if (purchaseOrderItems.getCode() != 0) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            JsonResult<List<OfflineInvoicingApplicationProjectMaterialPurchaseOrder>> listJsonResult = offlineInvoicingApplicationProjectMaterialPurchaseOrders(offlineInvoicingApplicationProjectMaterialPurchaseOrder);
            if (listJsonResult.getCode() != 0) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            List<ProjectSuppliesPurchaseOrderItem> projectSuppliesPurchaseOrderItems = purchaseOrderItems.getData();
            List<OfflineInvoicingApplicationProjectMaterialPurchaseOrder> offlineInvoicingApplicationProjectMaterialPurchaseOrders = listJsonResult.getData();
            log.info("项目物资采购订单:{}", JSON.toJSONString(projectSuppliesPurchaseOrderItems, true));
            log.info("项目物资采购订单线下开票申请表:{}", JSON.toJSONString(offlineInvoicingApplicationProjectMaterialPurchaseOrders, true));
            Valid valid = new Valid();
            List<ValidType> types = new ArrayList<>(2);
            valid.setTypes(types);

            ValidType type1 = new ValidType();
            type1.setSheetName("项目物资采购订单表");
            List<ValidResult> results1 = new ArrayList<>();
            type1.setResults(results1);
            types.add(type1);

            doValid(results1, N1, s -> s.anyMatch(x -> Objects.isNull(x.getNumber())), projectSuppliesPurchaseOrderItems.stream());
            doValid(results1, N2, s -> s.anyMatch(x -> StringUtils.isBlank(x.getPurchaseUnit())), projectSuppliesPurchaseOrderItems.stream());
            doValid(results1, N3, s -> s.anyMatch(x -> StringUtils.isBlank(x.getWsgs())), projectSuppliesPurchaseOrderItems.stream());
            doValid(results1, N4, s -> s.anyMatch(x -> StringUtils.isBlank(x.getQgr())), projectSuppliesPurchaseOrderItems.stream());
            doValid(results1, N5, s -> s.anyMatch(x -> Objects.isNull(x.getOrderTime())), projectSuppliesPurchaseOrderItems.stream());
            doValid(results1, N6, s -> s.anyMatch(x -> {
                if (Objects.isNull(x.getOrderMoney())) {
                    return true;
                }
                if (x.getOrderMoney() > 10000_0000D) {
                    return true;
                }
                String ms = x.getOrderMoney().toString();
                if (ms.contains(".") && (ms.length() - 1) - ms.indexOf('.') > 2) {
                    return true;
                } else if (ms.contains(".")) {
                    return false;
                } else {
                    return false;
                }
            }), projectSuppliesPurchaseOrderItems.stream());
            doValid(results1, N7, s -> s.anyMatch(x -> StringUtils.isBlank(x.getGysName())), projectSuppliesPurchaseOrderItems.stream());
            doValid(results1, N8, s -> s.anyMatch(x -> {
                if (StringUtils.isBlank(x.getFkms())) {
                    return true;
                }
                return !PAY_MODES.contains(x.getFkms());
            }), projectSuppliesPurchaseOrderItems.stream());
            doValid(results1, N9, s -> s.anyMatch(x -> {
                if (StringUtils.isBlank(x.getCglx())) {
                    return true;
                }
                return !PURCHASE_MODES.contains(x.getCglx());
            }), projectSuppliesPurchaseOrderItems.stream());
            doValid(results1, N10, s -> s.anyMatch(x -> {
                if (StringUtils.isBlank(x.getXmbh())) {
                    return false;
                }
                return x.getXmbh().split("、").length > 1 || x.getXmbh().split("，").length > 1 || x.getXmbh().split(",").length > 1;
            }), projectSuppliesPurchaseOrderItems.stream());
            doValid(results1, N11, s -> s.anyMatch(x -> {
                if (StringUtils.isBlank(x.getXmName())) {
                    return false;
                }
                return StringUtils.length(x.getXmName()) > 118;
            }), projectSuppliesPurchaseOrderItems.stream());
            doValid(results1, N12, s -> s.anyMatch(x -> {
                if (StringUtils.isBlank(x.getReceiver())) {
                    return true;
                }
                if (StringUtils.length(x.getReceiver()) > 10) {
                    return true;
                }
                return !validReceiver(x.getReceiver());
            }), projectSuppliesPurchaseOrderItems.stream());
            doValid(results1, N13, s -> s.anyMatch(x -> {
                String phone = x.getPhone();
                if (StringUtils.isBlank(phone)) {
                    return true;
                }
                return StringUtils.length(phone) != 11;
            }), projectSuppliesPurchaseOrderItems.stream());
            doValid(results1, N14, s -> s.anyMatch(x -> {
                String address = x.getAddress();
                if (StringUtils.isBlank(address)) {
                    return true;
                }
                return StringUtils.length(address) > 200;
            }), projectSuppliesPurchaseOrderItems.stream());
            doValid(results1, N15, s -> s.anyMatch(x -> {
                Integer yb = x.getYb();
                if (Objects.isNull(yb)) {
                    return true;
                }
                if (yb.toString().length() != 6) {
                    return true;
                }
                final String patternS = "[0-9]+";
                return !Pattern.matches(patternS, yb.toString());
            }), projectSuppliesPurchaseOrderItems.stream());
            doValid(results1, N16, s -> s.anyMatch(x -> {
                String mark = x.getMark();
                if (StringUtils.isBlank(mark)) {
                    return false;
                }
                return StringUtils.length(mark) > 200;
            }), projectSuppliesPurchaseOrderItems.stream());

            ValidType type2 = new ValidType();
            type2.setSheetName("项目物资采购订单线下开票申请表");
            List<ValidResult> results2 = new ArrayList<>();
            type2.setResults(results2);
            types.add(type2);

            Map<Integer, String> numberGysNameMap = projectSuppliesPurchaseOrderItems.stream().collect(Collectors.toMap(ProjectSuppliesPurchaseOrderItem::getNumber, ProjectSuppliesPurchaseOrderItem::getGysName));

            doValid2(results2, M1, s -> s.anyMatch(x -> {
                Integer fpxh = x.getFpxh();
                if (Objects.isNull(fpxh)) {
                    return true;
                }
                String gysmc = x.getGysmc();
                String beValid = numberGysNameMap.get(fpxh);
                return !StringUtils.equals(gysmc, beValid);
            }), offlineInvoicingApplicationProjectMaterialPurchaseOrders.stream());

            doValid2(results2, M2, s -> s.anyMatch(x -> {
                if (Objects.isNull(x.getOrderNum())) {
                    return true;
                }
                return !x.getFpxh().equals(x.getOrderNum());
            }), offlineInvoicingApplicationProjectMaterialPurchaseOrders.stream());

            doValid2(results2, M3, s -> s.anyMatch(x -> StringUtils.isBlank(x.getWsgs())), offlineInvoicingApplicationProjectMaterialPurchaseOrders.stream());
            doValid2(results2, M4, s -> s.anyMatch(x -> StringUtils.isBlank(x.getGhfmc())), offlineInvoicingApplicationProjectMaterialPurchaseOrders.stream());
            doValid2(results2, M5, s -> s.anyMatch(x -> StringUtils.isBlank(x.getSh()) || StringUtils.length(x.getSh()) != 18), offlineInvoicingApplicationProjectMaterialPurchaseOrders.stream());
            doValid2(results2, M6, s -> s.anyMatch(x -> {
                if (StringUtils.isBlank(x.getAddressPhone())) {
                    return true;
                }
                return !validAddressPhone(x.getAddressPhone());
            }), offlineInvoicingApplicationProjectMaterialPurchaseOrders.stream());
            doValid2(results2, M7, s -> s.anyMatch(x -> {
                if (StringUtils.isBlank(x.getYhzh())) {
                    return true;
                }
                return !validYhzh(x.getYhzh());
            }), offlineInvoicingApplicationProjectMaterialPurchaseOrders.stream());
            doValid2(results2, M8, s -> s.anyMatch(x -> StringUtils.isBlank(x.getPl()) || StringUtils.length(x.getPl()) > 20), offlineInvoicingApplicationProjectMaterialPurchaseOrders.stream());
            doValid2(results2, M9, s -> s.anyMatch(x -> StringUtils.isBlank(x.getWlmc()) || StringUtils.length(x.getWlmc()) > 20), offlineInvoicingApplicationProjectMaterialPurchaseOrders.stream());
            doValid2(results2, M10, s -> s.anyMatch(x -> StringUtils.isBlank(x.getSpmc()) || StringUtils.length(x.getSpmc()) > 45), offlineInvoicingApplicationProjectMaterialPurchaseOrders.stream());
            doValid2(results2, M11, s -> s.anyMatch(x -> StringUtils.isBlank(x.getGgxh()) || StringUtils.length(x.getGgxh()) > 20), offlineInvoicingApplicationProjectMaterialPurchaseOrders.stream());
            doValid2(results2, M12, s -> s.anyMatch(x -> StringUtils.isBlank(x.getUnit())), offlineInvoicingApplicationProjectMaterialPurchaseOrders.stream());
            doValid2(results2, M13, s -> s.anyMatch(x -> Objects.isNull(x.getNumber())), offlineInvoicingApplicationProjectMaterialPurchaseOrders.stream());
            doValid2(results2, M14, s -> s.anyMatch(x -> Objects.isNull(x.getDj())), offlineInvoicingApplicationProjectMaterialPurchaseOrders.stream());
            doValid2(results2, M15, s -> s.anyMatch(x -> Objects.isNull(x.getSl())), offlineInvoicingApplicationProjectMaterialPurchaseOrders.stream());
            doValid2(results2, M16, s -> s.anyMatch(x -> Objects.isNull(x.getMoney())), offlineInvoicingApplicationProjectMaterialPurchaseOrders.stream());
            doValid2(results2, M17, s -> s.anyMatch(x -> StringUtils.isBlank(x.getSpr())), offlineInvoicingApplicationProjectMaterialPurchaseOrders.stream());
            doValid2(results2, M18, s -> s.anyMatch(x -> StringUtils.isBlank(x.getSprgsmc())), offlineInvoicingApplicationProjectMaterialPurchaseOrders.stream());
            doValid2(results2, M19, s -> s.anyMatch(x -> StringUtils.isBlank(x.getLxdh())), offlineInvoicingApplicationProjectMaterialPurchaseOrders.stream());
            doValid2(results2, M20, s -> s.anyMatch(x -> StringUtils.isBlank(x.getAddress())), offlineInvoicingApplicationProjectMaterialPurchaseOrders.stream());
            doValid2(results2, M21, s -> s.anyMatch(x -> StringUtils.isBlank(x.getGysmc())), offlineInvoicingApplicationProjectMaterialPurchaseOrders.stream());
            doValid2(results2, M22, s -> s.anyMatch(x -> StringUtils.isBlank(x.getGyslxr())), offlineInvoicingApplicationProjectMaterialPurchaseOrders.stream());
            doValid2(results2, M23, s -> s.anyMatch(x -> StringUtils.isBlank(x.getGyslxdh())), offlineInvoicingApplicationProjectMaterialPurchaseOrders.stream());
            doValid2(results2, M24, s -> s.anyMatch(x -> StringUtils.isBlank(x.getGysdzyj())), offlineInvoicingApplicationProjectMaterialPurchaseOrders.stream());
            doValid2(results2, M25, s -> s.anyMatch(x -> StringUtils.isBlank(x.getTzryyx())), offlineInvoicingApplicationProjectMaterialPurchaseOrders.stream());
            doValid2(results2, M26, s -> s.anyMatch(x -> Objects.isNull(x.getGyszkl())), offlineInvoicingApplicationProjectMaterialPurchaseOrders.stream());
            doValid2(results2, M27, s -> s.anyMatch(x -> Objects.isNull(x.getCbj())), offlineInvoicingApplicationProjectMaterialPurchaseOrders.stream());
            doValid2(results2, M28, s -> s.anyMatch(x -> Objects.isNull(x.getGysydTime())), offlineInvoicingApplicationProjectMaterialPurchaseOrders.stream());
            doValid2(results2, M29, s -> s.anyMatch(x -> Objects.isNull(x.getJxpjdydTime())), offlineInvoicingApplicationProjectMaterialPurchaseOrders.stream());
            Map<String, String> lxrMapDh = new HashMap<>();
            Map<String, String> lxrMapYj = new HashMap<>();
            for (OfflineInvoicingApplicationProjectMaterialPurchaseOrder order : offlineInvoicingApplicationProjectMaterialPurchaseOrders) {
                String gyslxr = order.getGyslxr();
                if (!lxrMapDh.containsKey(gyslxr)) {
                    lxrMapDh.put(gyslxr, order.getGyslxdh());
                }
                if (!lxrMapYj.containsKey(gyslxr)) {
                    lxrMapYj.put(gyslxr, order.getGysdzyj());
                }
            }
            doValid2(results2, M30, s -> s.anyMatch(x -> {
                String gyslxr = x.getGyslxr();
                String gyslxdh = x.getGyslxdh();
                String gysdzyj = x.getGysdzyj();
                return !StringUtils.equals(lxrMapDh.get(gyslxr), gyslxdh) || !StringUtils.equals(lxrMapYj.get(gyslxr), gysdzyj);
            }), offlineInvoicingApplicationProjectMaterialPurchaseOrders.stream());
            doValid2(results2, M31, s -> s.anyMatch(x -> !StringUtils.equals(x.getGhfmc(), x.getSprgsmc())), offlineInvoicingApplicationProjectMaterialPurchaseOrders.stream());
            return new JsonResult<>(0, "", valid);
        } catch (Exception e) {
            log.error("解析Excel发生异常:", e);
            return new JsonResult<>(3, e.getMessage(), null);
        }
    }

    private static void doValid(List<ValidResult> results, String valid, Predicate<Stream<ProjectSuppliesPurchaseOrderItem>> predicate, Stream<ProjectSuppliesPurchaseOrderItem> stream) {
        ValidResult of = ValidResult.of(valid);
        if (predicate.test(stream)) {
            of.setIsTrue(false);
        }
        results.add(of);
    }

    private static void doValid2(List<ValidResult> results, String valid, Predicate<Stream<OfflineInvoicingApplicationProjectMaterialPurchaseOrder>> predicate, Stream<OfflineInvoicingApplicationProjectMaterialPurchaseOrder> stream) {
        ValidResult of = ValidResult.of(valid);
        if (predicate.test(stream)) {
            of.setIsTrue(false);
        }
        results.add(of);
    }

    private static JsonResult<List<OfflineInvoicingApplicationProjectMaterialPurchaseOrder>> offlineInvoicingApplicationProjectMaterialPurchaseOrders(XSSFSheet offlineInvoicingApplicationProjectMaterialPurchaseOrder) {
        List<OfflineInvoicingApplicationProjectMaterialPurchaseOrder> orders = new ArrayList<>();
        final int start = 4;
        // 获得数据的总行数
        int lastRowNum = offlineInvoicingApplicationProjectMaterialPurchaseOrder.getLastRowNum();
        for (int i = start; i < lastRowNum; i++) {
            // 获得第i行对象
            Row row = offlineInvoicingApplicationProjectMaterialPurchaseOrder.getRow(i);
            if (row == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            OfflineInvoicingApplicationProjectMaterialPurchaseOrder order = new OfflineInvoicingApplicationProjectMaterialPurchaseOrder();
            Cell cell = row.getCell(1);
            if (cell != null) {
                order.setFplx(cell.getStringCellValue());
            }
            cell = row.getCell(2);
            if (cell == null) {
                break;
            }
            String s = parseNumber(cell);
            if (StringUtils.isBlank(s)) {
                break;
            }
            order.setFpxh(Double.valueOf(s).intValue());
            cell = row.getCell(3);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            order.setOrderNum(Double.valueOf(parseNumber(cell)).intValue());
            cell = row.getCell(4);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            order.setWsgs(cell.getStringCellValue());
            cell = row.getCell(5);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            order.setGhfmc(cell.getStringCellValue());
            cell = row.getCell(6);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            order.setSh(cell.getStringCellValue());
            cell = row.getCell(7);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            order.setAddressPhone(cell.getStringCellValue());
            cell = row.getCell(8);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            order.setYhzh(cell.getStringCellValue());
            cell = row.getCell(9);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            order.setPl(cell.getStringCellValue());
            cell = row.getCell(10);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            order.setWlmc(cell.getStringCellValue());
            cell = row.getCell(11);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            order.setSpmc(cell.getStringCellValue());
            cell = row.getCell(12);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            order.setGgxh(parseNumber(cell));
            cell = row.getCell(13);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            order.setUnit(cell.getStringCellValue());
            cell = row.getCell(14);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            order.setNumber(Double.parseDouble(parseNumber(cell)));
            cell = row.getCell(15);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            order.setDj(Double.parseDouble(parseNumber(cell)));
            cell = row.getCell(16);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            order.setSl(Double.parseDouble(parseNumber(cell)));
            cell = row.getCell(17);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            order.setMoney(Double.parseDouble(parseNumber(cell)));
            cell = row.getCell(18);
            if (cell != null) {
                order.setMark(cell.getStringCellValue());
            }
            cell = row.getCell(19);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            order.setSpr(cell.getStringCellValue());
            cell = row.getCell(20);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            order.setSprgsmc(cell.getStringCellValue());
            cell = row.getCell(21);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            order.setLxdh(parseNumber(cell));
            cell = row.getCell(22);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            order.setAddress(cell.getStringCellValue());
            cell = row.getCell(23);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            order.setGysmc(cell.getStringCellValue());
            cell = row.getCell(24);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            order.setGyslxr(cell.getStringCellValue());
            cell = row.getCell(25);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            order.setGyslxdh(parseNumber(cell));
            cell = row.getCell(26);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            order.setGysdzyj(cell.getStringCellValue());
            cell = row.getCell(27);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            order.setTzryyx(cell.getStringCellValue());
            cell = row.getCell(28);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            order.setGyszkl(Double.parseDouble(parseNumber(cell)));
            cell = row.getCell(29);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            order.setCbj(Double.parseDouble(parseNumber(cell)));
            cell = row.getCell(30);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            order.setGysydTime(parseDate(cell));
            cell = row.getCell(31);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            order.setJxpjdydTime(parseDate(cell));
            orders.add(order);
        }
        return new JsonResult<>(0, "", orders);
    }

    private static JsonResult<List<ProjectSuppliesPurchaseOrderItem>> projectSuppliesPurchaseOrderItems(XSSFSheet projectSuppliesPurchaseOrderTemplate) {
        List<ProjectSuppliesPurchaseOrderItem> items = new ArrayList<>();
        final int start = 4;
        // 获得数据的总行数
        int lastRowNum = projectSuppliesPurchaseOrderTemplate.getLastRowNum();
        for (int i = start; i < lastRowNum; i++) {
            // 获得第i行对象
            Row row = projectSuppliesPurchaseOrderTemplate.getRow(i);
            if (row == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            Cell cell = row.getCell(1);
            if (cell == null) {
                break;
            }
            ProjectSuppliesPurchaseOrderItem item = new ProjectSuppliesPurchaseOrderItem();
            item.setNumber(Double.valueOf(parseNumber(cell)).intValue());
            cell = row.getCell(2);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            item.setPurchaseUnit(cell.getStringCellValue());
            cell = row.getCell(3);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            item.setWsgs(cell.getStringCellValue());
            cell = row.getCell(4);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            item.setQgr(cell.getStringCellValue());
            cell = row.getCell(5);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            item.setOrderTime(cell.getDateCellValue());
            cell = row.getCell(6);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            item.setOrderMoney(Double.parseDouble(parseNumber(cell)));
            cell = row.getCell(7);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            item.setGysName(cell.getStringCellValue());
            cell = row.getCell(8);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            item.setFkms(cell.getStringCellValue());
            cell = row.getCell(9);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            item.setCglx(cell.getStringCellValue());
            cell = row.getCell(10);
            if (cell != null) {
                item.setXmbh(cell.getStringCellValue());
            }
            cell = row.getCell(11);
            if (cell != null) {
                item.setXmName(cell.getStringCellValue());
            }
            cell = row.getCell(12);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            item.setReceiver(cell.getStringCellValue());
            cell = row.getCell(13);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            item.setPhone(parseNumber(cell));
            cell = row.getCell(14);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            item.setAddress(cell.getStringCellValue());
            cell = row.getCell(15);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            item.setYb(Double.valueOf(parseNumber(cell)).intValue());
            cell = row.getCell(16);
            if (cell != null) {
                item.setMark(cell.getStringCellValue());
            }
            items.add(item);
        }
        return new JsonResult<>(0, "", items);
    }

    private static String parseNumber(Cell cell) {
        if (cell.getCellType() == CellType.NUMERIC) {
            if (!String.valueOf(cell.getNumericCellValue()).contains("E")) {
                return String.valueOf(cell.getNumericCellValue());
            } else {
                return new DecimalFormat("#").format(cell.getNumericCellValue());
            }
        } else if (cell.getCellType() == CellType.FORMULA) {
            FormulaEvaluator formulaEvaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
            CellValue evaluate = formulaEvaluator.evaluate(cell);
            return evaluate.formatAsString();
        } else {
            return cell.getStringCellValue();
        }
    }

    private static Date parseDate(Cell cell) {
        try {
             return cell.getDateCellValue();
        } catch (Exception e) {
            return null;
        }
    }

    private static boolean validReceiver(String receiver) {
        final String patternH = "[\\u4E00-\\u9FA5]+";
        final String patternS = "[0-9]+";

        for (int i = 0; i < receiver.length(); i++) {
            CharSequence charSequence = receiver.subSequence(i, i + 1);
            if (!Pattern.matches(patternH, charSequence) && !Pattern.matches(patternS, charSequence)) {
                return false;
            }
        }

        return true;
    }

    private static boolean validYhzh(String yhzh) {
        yhzh = yhzh.replace(" ", "");
        final String patternH = "[\\u4e00-\\u9fa5_a-zA-Z0-9]+[0-9]{10,}";
        return Pattern.matches(patternH, yhzh);
    }

    private static boolean validAddressPhone(String addressPhone) {
        addressPhone = addressPhone.replace(" ", "");
        addressPhone = addressPhone.replace("-", "");
        addressPhone = addressPhone.replace("－", "");
        addressPhone = addressPhone.replace("—", "");
        final String a = "[\\u4e00-\\u9fa5_a-zA-Z0-9]+[((13[0-9])|(14[0-9])|(15[0-9])|(16[0-9])|(17[0-9])|(18[0-9])|(19[0-9]))\\d{8}|1\\d{10}|0\\d{9,11}]{4,}";
        final String b = "[((13[0-9])|(14[0-9])|(15[0-9])|(16[0-9])|(17[0-9])|(18[0-9])|(19[0-9]))\\d{8}|1\\d{10}|0\\d{9,11}]{4,}+[\\u4e00-\\u9fa5_a-zA-Z0-9]+";
        return Pattern.matches(a, addressPhone) || Pattern.matches(b, addressPhone);
    }
}
