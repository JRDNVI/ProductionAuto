# ProductionAuto

An algorithm that will assess stock, depot days, orders, etc., to create a production plan.

## Workflow Overview

```mermaid
flowchart TD
  subgraph PowerQueryStaging["Power Query Staging"]
    direction LR
    PQ_Orders["PQ_Orders\n(Load Orders CSVs)"]
    PQ_Master["PQ_MasterData\n(Load Product Codes Master)"]
    PQ_Stocks["PQ_Stocks\n(Load Live Stocks)"]
    PQ_NetReq["PQ_NetReq\n(Compute Net Requirements)"]
    PQ_RawMat["PQ_RawMaterials\n(Aggregate Raw Mtrl Pull)"]
    PQ_Storage["PQ_Storage\n(Optional: Storage Dept Data)"]
  end

  Start(("Start: BuildProductionPlan"))
  Refresh(("Refresh Power Query Connections"))
  Allocate(("DistributeAcrossLines\n(Creates 'Allocations' sheet)"))
  GenLines["GenerateLineSheets\n(Line‑specific tabs)"]
  UpdateRM["UpdateRawMaterialSheet\n(Raw Material Daily Requirement)"]
  UpdateStore["UpdateStorageSheet\n(Equaliser/Storage sheet)"]
  End(("Done"))

  Start --> Refresh
  Refresh --> PQ_Orders
  Refresh --> PQ_Master
  Refresh --> PQ_Stocks
  Refresh --> PQ_RawMat
  Refresh --> PQ_Storage

  PQ_Orders --> PQ_NetReq
  PQ_Master --> PQ_NetReq
  PQ_Stocks --> PQ_NetReq

  PQ_NetReq --> Allocate
  Allocate --> GenLines
  PQ_RawMat --> UpdateRM
  PQ_Storage --> UpdateStore

  GenLines --> End
  UpdateRM --> End
  UpdateStore --> End
  ```
