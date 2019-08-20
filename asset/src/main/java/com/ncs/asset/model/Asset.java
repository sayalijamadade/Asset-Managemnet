package com.ncs.asset.model;

import lombok.*;

import javax.persistence.*;
import java.util.Date;
import java.util.Set;

@Entity
@Data
@NoArgsConstructor
public class Asset {

    @Id
    @GeneratedValue
    private int assetId;

    @Column
    private String typeName;

    @Column
    private String groupName;

    @Column
    private String assetName;

    @Column
    private String vendorName;

    @Column
    private Long cost;

    @Column
    private String description;

    @OneToOne(mappedBy = "asset", cascade = CascadeType.ALL)
    private AssetStatus assetStatus;

    @ManyToOne
    private Location location;

}
