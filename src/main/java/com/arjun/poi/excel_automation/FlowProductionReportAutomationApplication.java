package com.arjun.poi.excel_automation;

import com.arjun.poi.excel_automation.service.ServiceClass;
import lombok.extern.slf4j.Slf4j;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.ApplicationContext;
import org.springframework.context.annotation.ComponentScan;
import org.springframework.data.jpa.repository.config.EnableJpaAuditing;

import java.io.IOException;

@SpringBootApplication
@Slf4j
@ComponentScan("com.ad2pro.spectra")
@EnableJpaAuditing

public class FlowProductionReportAutomationApplication {

	public static void main(String[] args) throws IOException {
		ApplicationContext applicationContext = SpringApplication.run(FlowProductionReportAutomationApplication.class, args);
		ServiceClass serviceClass = applicationContext.getBean(ServiceClass.class);
		serviceClass.generateExcel();
	}

}
