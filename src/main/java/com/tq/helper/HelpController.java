package com.tq.helper;

import lombok.extern.slf4j.Slf4j;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

/**
 * @author HuSen
 * create on 2019/12/11 11:19
 */
@Slf4j
@RestController
@RequestMapping("/api/help")
public class HelpController {

    @PostMapping
    public JsonResult<Valid> valid(MultipartFile file) {
        try {
            if (file == null) {
                return new JsonResult<>(5, "文件不能为空", null);
            }
            return HelpUtils.doHelp(file.getInputStream());
        } catch (Exception e) {
            log.error("获取上文件流失败:", e);
            return new JsonResult<>(6, "获取上文件流失败", null);
        }
    }

}
