package com.ncs.asset.dto;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
public class AssetDto {

    private String typeName;

    private String groupName;

    private String assetName;

    private String vendorName;

    private Long cost;

    private String description;
}
