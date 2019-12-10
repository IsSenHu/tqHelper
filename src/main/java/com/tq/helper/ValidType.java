package com.tq.helper;

import lombok.Data;

import java.util.List;

/**
 * @author HuSen
 * create on 2019/12/10
 */
@Data
class ValidType {

    private String sheetName;

    private List<ValidResult> results;
}
