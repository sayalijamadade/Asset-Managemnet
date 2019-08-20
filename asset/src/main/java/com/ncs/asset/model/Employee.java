package com.ncs.asset.model;

import lombok.*;

import javax.persistence.*;
import java.util.Set;

@Entity
@Data
@NoArgsConstructor
public class Employee {

    @Id
    @GeneratedValue
    private int empId;

    @Column
    private String empName;

    @OneToMany(cascade=CascadeType.ALL)
    @JoinColumn(name="transferId")
    private Set<AssetTransfer> assetTransferSet;
}
