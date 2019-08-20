package com.ncs.asset.services.implementation;

import com.ncs.asset.model.Location;
import com.ncs.asset.repository.LocationRepository;
import com.ncs.asset.services.interfaces.LocationService;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

@Service
public class LocationServiceImp implements LocationService {

    @Autowired
    LocationRepository locationRepository;

    @Override
    public Location saveLocation(Location location) {
        return locationRepository.save(location) ;
    }
}
