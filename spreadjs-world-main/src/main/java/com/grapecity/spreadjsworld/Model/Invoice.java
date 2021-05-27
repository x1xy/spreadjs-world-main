package com.grapecity.spreadjsworld.Model;

import java.util.Date;
import java.util.List;

public class Invoice {

    public Company company ;
    public String number ;
    public Date date ;
    public Customer customer ;
    public Customer receiverCustomer ;
    public List<Record> records ;


    public Invoice(Company company,String number, Date date, Customer customer, Customer receiverCustomer, List<Record> records) {
        this.company = company;
        this.number = number;
        this.date = date;
        this.customer = customer;
        this.receiverCustomer = receiverCustomer;
        this.records = records;
    }
}
