package com.ncs.asset.repository;

import com.ncs.asset.model.AssetTransfer;
import com.ncs.asset.model.Employee;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface AssetTransferRepository extends JpaRepository<AssetTransfer, Integer> {
}
