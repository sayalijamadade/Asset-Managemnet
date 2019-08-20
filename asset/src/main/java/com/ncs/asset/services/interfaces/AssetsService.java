package com.ncs.asset.services.interfaces;

import com.ncs.asset.dto.AssetDto;
import com.ncs.asset.model.Asset;

import java.util.List;

public interface AssetsService {

    Asset saveAsset(AssetDto assetDto, String location);

    List<Asset> fetchAssets();
}
