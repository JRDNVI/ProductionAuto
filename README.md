# ProductionAuto
An algorithm that will assess stock, depo days, orders, etc.. to create a production plan.

## Workflow Overview

```mermaid
flowchart TD
  subgraph subGraph0["Power Query Staging"]
    direction LR
      PQ_Orders(["PQ_Orders(Load Orders CSVs)"])
      PQ_Master(["PQ_MasterData(Load Product Codes Master)"])
      PQ_Stocks(["PQ_Stocks(Load Live Stocks)"])
      PQ_NetReq(["PQ_NetReq(Compute Net Requirements)"])
      PQ_RawMat(["PQ_RawMaterials(Aggregate Raw Mtrl Pull)"])
      PQ_Storage(["PQ_Storage(Optional: Storage Dept Data)"])
  end

  Start(["Start: BuildProductionPlan"])
  Refresh(["Refresh Power Query Connections"])
  Allocate{ "DistributeAcrossLines\n(Creates 'Allocations' sheet)" }

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
  Allocate --> GenLines(["GenerateLineSheets\n(Line-specific tabs)"])
  PQ_RawMat --> UpdateRM(["UpdateRawMaterialSheet\n(Raw Material Daily Requirement)"])
  PQ_Storage --> UpdateStore(["UpdateStorageSheet\n(Equaliser/Storage sheet)"])

  GenLines --> End(["Done"])
  UpdateRM --> End
  UpdateStore --> End
  ```