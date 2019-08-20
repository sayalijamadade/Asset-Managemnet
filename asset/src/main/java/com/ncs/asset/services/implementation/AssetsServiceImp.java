package com.ncs.asset.services.implementation;

import com.ncs.asset.dto.AssetDto;
import com.ncs.asset.model.Asset;
import com.ncs.asset.model.AssetStatus;
import com.ncs.asset.model.Location;
import com.ncs.asset.repository.AssetsRepository;
import com.ncs.asset.repository.LocationRepository;
import com.ncs.asset.services.interfaces.AssetsService;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;

@Service
public class AssetsServiceImp implements AssetsService {

    @Autowired
    AssetsRepository assetsRepository;



    @Autowired
    LocationRepository locationRepository;

    @Override
    public Asset saveAsset(AssetDto assetDto, String location) {

        Location location1 = locationRepository.findByLocationName(location);

        Asset asset = new Asset();
        asset.setAssetName(assetDto.getAssetName());
        asset.setCost(assetDto.getCost());
        asset.setDescription(assetDto.getDescription());
        asset.setGroupName(assetDto.getGroupName());
        asset.setLocation(location1);
        asset.setTypeName(assetDto.getTypeName());
        asset.setVendorName(assetDto.getVendorName());
        //asset.setAssetStatus(new AssetStatus());

        return assetsRepository.save(asset);
    }

    @Override
    public List<Asset> fetchAssets() {
        return assetsRepository.findAll();

    }
}
