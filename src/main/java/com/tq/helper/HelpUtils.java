package com.tq.helper;

import com.alibaba.fastjson.JSON;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.util.*;
import java.util.function.Predicate;
import java.util.regex.Pattern;
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
    private static final String M2 = "订单号必填";
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
    private static final String M28 = "供应商与英大合同的提供必填";
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
            order.setGysydTime(cell.getDateCellValue());
            cell = row.getCell(31);
            if (cell == null) {
                return new JsonResult<>(2, "请检查结构是否正确", null);
            }
            order.setJxpjdydTime(cell.getDateCellValue());
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

    public static void main(String[] args) throws Exception {
        InputStream inputStream = new FileInputStream(new File("D:\\附件1模板(绵阳2019.12.10.xlsx"));
        System.out.println(doHelp(inputStream));
    }
}
