package com.ncs.asset.controller;


import com.ncs.asset.model.Location;
import com.ncs.asset.services.interfaces.LocationService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

import java.util.List;

@RestController
@RequestMapping(value = "/location")
public class LocationController {

    @Autowired
    LocationService locationService;

    @PostMapping(value = "/save")
    public Location persistLocation(@RequestBody final Location location) {

        return locationService.saveLocation(location);
    }

}
