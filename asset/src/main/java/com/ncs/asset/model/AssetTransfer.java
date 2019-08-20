package com.ncs.asset.model;


import lombok.Data;
import lombok.NoArgsConstructor;

import javax.persistence.*;
import java.util.Date;

@Entity
@Data
@NoArgsConstructor
public class AssetTransfer {

    @Id
    @GeneratedValue
    private int transferId;

    @Column
    private String transferFrom;

    @Column
    private String transferTo;

    @Column
    private Date transferDate;

    @Column
    private Date returnDate;

    @ManyToOne
    private Employee employee;
}
