package com.example.dynamic.map.entity;

import jakarta.persistence.*;

import java.util.HashMap;
import java.util.Map;

@Entity
@Table(name = "basic")
public class Basic {
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private int id;

    @Column(name = "company_pan_no")
    private String panNo;

    @Column(name = "age")
    private String age;

    @ElementCollection
    @CollectionTable(name = "basic_dynamic_fields", joinColumns = @JoinColumn(name = "basic_id"))
    @MapKeyColumn(name = "field_name")
    @Column(name = "field_value")
    private Map<String, String> dynamicFields = new HashMap<>();

    public void setDynamicField(String fieldName, String value) {
        dynamicFields.put(fieldName, value);
    }

    public String getDynamicField(String fieldName) {
        return dynamicFields.get(fieldName);
    }
}
