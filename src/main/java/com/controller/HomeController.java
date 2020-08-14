package com.controller;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

@Controller
public class HomeController {
    Logger logger = LoggerFactory.getLogger(HomeController.class);
    private Integer hitCount = 0;
    @GetMapping("loginLogOffHome")
    public String homePage() {
        logger.info("Number of hits: {}", ++hitCount);
        return "index";
    }
}
