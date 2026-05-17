package com.yennor;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.scheduling.annotation.EnableScheduling;

@SpringBootApplication
@EnableScheduling
public class YennorApplication {

    public static void main(String[] args) {
        SpringApplication.run(YennorApplication.class, args);
    }
}
