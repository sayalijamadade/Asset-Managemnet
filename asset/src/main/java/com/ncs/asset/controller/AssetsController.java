package com.ncs.asset.controller;

import com.ncs.asset.dto.AssetDto;
import com.ncs.asset.model.Asset;
import com.ncs.asset.model.Location;
import com.ncs.asset.services.interfaces.AssetsService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

import java.util.List;

@RestController
@RequestMapping(value = "/assets")
public class AssetsController {

    @Autowired
    AssetsService assetsService;

    @PostMapping(value = "/save")
    public Asset persistAsset(@RequestBody final AssetDto assetDto, @RequestParam("location") String location) {

        return assetsService.saveAsset(assetDto, location);
    }


    @GetMapping(value = "/")
    public List<Asset> getAll()
    {
        return assetsService.fetchAssets();
    }
}
