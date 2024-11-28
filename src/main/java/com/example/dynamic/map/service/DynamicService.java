package com.example.dynamic.map.service;

import com.example.dynamic.map.entity.Basic;
import jakarta.persistence.EntityManager;
import jakarta.transaction.Transactional;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

@Service
public class DynamicService {
    @Autowired
    private EntityManager entityManager;

    public DynamicService(EntityManager entityManager) {
        this.entityManager = entityManager;
    }

    @Transactional
    public void addDynamicField(String fieldName, String fieldType) {
        String sql = String.format("ALTER TABLE account_details_dynamic_fields ADD COLUMN %s %s", fieldName, fieldType);
        entityManager.createNativeQuery(sql).executeUpdate();
    }

    @Transactional
    public void setDynamicFieldValue(Long accountDetailsId, String fieldName, String value) {
        Basic accountDetails = entityManager.find(Basic.class, accountDetailsId);
        if (accountDetails != null) {
            accountDetails.setDynamicField(fieldName, value);
            entityManager.merge(accountDetails);
        }
    }

    public String getDynamicFieldValue(Long accountDetailsId, String fieldName) {
        Basic accountDetails = entityManager.find(Basic.class, accountDetailsId);
        if (accountDetails != null) {
            return accountDetails.getDynamicField(fieldName);
        }
        return null;
    }
}
