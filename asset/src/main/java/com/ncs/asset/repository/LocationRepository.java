package com.ncs.asset.repository;

import com.ncs.asset.model.Location;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;
import org.springframework.stereotype.Repository;

@Repository
public interface LocationRepository extends JpaRepository<Location, Integer> {

    @Query(
            value = "SELECT * FROM Location l WHERE l.loc_name = :location",
            nativeQuery = true)
    Location findByLocationName(@Param("location") String location);
}
