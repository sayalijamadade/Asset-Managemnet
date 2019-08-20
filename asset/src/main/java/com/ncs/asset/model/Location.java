package com.ncs.asset.model;

import lombok.*;

import javax.persistence.*;
import java.util.Set;

@Entity
@Data
@NoArgsConstructor
public class Location {

    @Id
    @GeneratedValue
    private int locId;

    @Column
    private String locName;


    @OneToMany(cascade=CascadeType.ALL)
    @JoinColumn(name="assetId")
    private Set<Asset> assets;


}
