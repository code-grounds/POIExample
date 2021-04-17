package com.arjun.poi.excel_automation.repository;
import com.arjun.poi.excel_automation.dao.RepositoryName;
import org.springframework.stereotype.Repository;
import org.springframework.data.repository.CrudRepository;

import java.util.List;

@Repository
public interface ITableNameRepository extends CrudRepository<RepositoryName, Integer> {

//    methods other then exposed by library goes here
//    datatype<TableName> findByColumnName(datatype columnVariableNme);
}