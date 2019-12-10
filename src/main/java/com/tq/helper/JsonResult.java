package com.tq.helper;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.Serializable;

/**
 * @author HuSen
 * create on 2019/7/15 11:01
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
class JsonResult<T> implements Serializable {

    private Integer code;

    private String message;

    private T data;
}
