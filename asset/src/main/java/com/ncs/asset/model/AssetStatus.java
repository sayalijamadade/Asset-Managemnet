package com.ncs.asset.model;


import lombok.Data;
import lombok.NoArgsConstructor;

import javax.persistence.*;
import java.util.Date;

@Entity
@Data
@NoArgsConstructor
public class AssetStatus {

    @Id
    @GeneratedValue
    private int discardedId;

    @Column
    private boolean isDiscarded;

    @Column
    private Date discardDate;

    @Column
    private String reasonForDiscard;

    @OneToOne
    @MapsId
    private Asset asset;
}
