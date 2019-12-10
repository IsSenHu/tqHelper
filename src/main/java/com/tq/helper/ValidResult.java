package com.tq.helper;

import lombok.AllArgsConstructor;
import lombok.Data;

/**
 * @author HuSen
 * create on 2019/12/10
 */
@Data
@AllArgsConstructor
class ValidResult {
    private String valid;
    private Boolean isTrue;

    static ValidResult of(String valid) {
        return new ValidResult(valid, true);
    }
}
