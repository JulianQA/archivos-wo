const archiveInput = document.querySelector(".form__archive");

archiveInput.addEventListener("change", async () => {
  const response = await readXlsxFile(archiveInput.files[0]);
  const nameArchiveRead = archiveInput.files[0].name;
  const name = nameArchiveRead.split(" ");
  const woName = `ARCHIVO WO ${name[0]} ${name[2]} ${name[3]} ${
    name[4].split(".")[0]
  }`;
  const excel = new Excel(response, name).getRow();
  excel.unshift([
    {
      value: "TipoDoc",
      fontWeight: "bold",
    },
    {
      value: "NumDoc",
      fontWeight: "bold",
    },
    {
      value: "FerchaElaboracion",
      fontWeight: "bold",
    },
    {
      value: "siglaMoneda",
      fontWeight: "bold",
    },
    {
      value: "tasaDeCambio",
      fontWeight: "bold",
    },
    {
      value: "vCuentaContable",
      fontWeight: "bold",
    },
    {
      value: "vNit",
      fontWeight: "bold",
    },
    {
      value: "Sucursal",
      fontWeight: "bold",
    },
    {
      value: "CodProducto",
      fontWeight: "bold",
    },
    {
      value: "Bodega",
      fontWeight: "bold",
    },
    {
      value: "Accion",
      fontWeight: "bold",
    },
    {
      value: "CantidadProducto",
      fontWeight: "bold",
    },
    {
      value: "Prefijo",
      fontWeight: "bold",
    },
    {
      value: "Consecutivo",
      fontWeight: "bold",
    },
    {
      value: "NumCuota",
      fontWeight: "bold",
    },
    {
      value: "FechaVencimiento",
      fontWeight: "bold",
    },
    {
      value: "CodImpuesto",
      fontWeight: "bold",
    },
    {
      value: "CodGrupoActivoFijo",
      fontWeight: "bold",
    },
    {
      value: "CodActivoFijo",
      fontWeight: "bold",
    },
    {
      value: "Descripicion",
      fontWeight: "bold",
    },
    {
      value: "CodSubCentroCostos",
      fontWeight: "bold",
    },
    {
      value: "Debito",
      fontWeight: "bold",
    },
    {
      value: "Credito",
      fontWeight: "bold",
    },
    {
      value: "Observaciones",
      fontWeight: "bold",
    },
    {
      value: "BaseGravable",
      fontWeight: "bold",
    },
    {
      value: "BaseExcenta",
      fontWeight: "bold",
    },
    {
      value: "MesCierre",
      fontWeight: "bold",
    },
  ]);
  return await writeXlsxFile(excel, {
    fileName: woName,
  });
});

class Excel {
  constructor(archive, name) {
    this.archive = archive;
    this.archiveName = name[0];
  }

  getRow() {
    let rows = [];
    for (let i = 1; i < this.archive.length; i++) {
      rows.push(Object.values(new Rows(this.archive[i], this.archiveName)));
    }
    return rows;
  }
}

class Rows {
  constructor(rows, archiveName) {
    this.tipoDoc = { type: String, value: "FV" };
    this.numDoc = { type: Number, value: rows[1] };
    this.fechaElaboracion = { type: String, value: rows[2] };
    this.siglaMoneda = { type: String, value: null };
    this.tasaDeCambio = { type: String, value: null };
    this.vCuentaContable = {
      type: String,
      value: this.validation(
        rows[19],
        archiveName,
        rows[5] == null ? " " : rows[5]
      ),
    };
    this.vNit = { type: Number, value: 222222222 };
    this.sucursal = { type: String, value: null };
    this.codProducto = { type: String, value: rows[8] };
    this.bodega = { type: String, value: null };
    this.accion = { type: String, value: null };
    this.cantidadProducto = { type: String, value: rows[11] };
    this.prefijo = { type: String, value: null };
    this.consecutivo = { type: String, value: null };
    this.numCuota = { type: String, value: null };
    this.fechaVencimiento = { type: String, value: null };
    this.codImpuesto = { type: String, value: rows[16] };
    this.codGrupoActivoFijo = { type: String, value: null };
    this.codActivoFijo = { type: String, value: null };
    this.descripcion = {
      type: String,
      value: this.validation(rows[19], archiveName),
    };
    this.codSubCentrocostos = { type: String, value: null };
    this.debito = { type: Number, value: rows[21] };
    this.credito = { type: Number, value: rows[22] };
    this.observaciones = { type: String, value: null };
    this.baseGravable = { type: String, value: null };
    this.baseExcenta = { type: String, value: null };
    this.mesCierre = { type: String, value: null };
  }
  validation(descripcion, archiveName, defaultValue = undefined) {
    if (descripcion === "TCREDITO" || descripcion === "TDEBITO") {
      if (archiveName === "CHIA") {
        return defaultValue ? "13359004" : "Bold";
      } else {
        return defaultValue ? "13359003" : "Credibanco";
      }
    } else if (descripcion === "DAVIVIENDA 3959") {
      return defaultValue ? "11200505" : "Davivienda 3959";
    } else if (descripcion === "WOMPI") {
      return defaultValue ? "13359002" : "Wompi";
    } else if (descripcion === "Transferencia bancar") {
      switch (archiveName) {
        case "CHIA":
          return defaultValue ? "11200503" : "Bancolombia 1663 Chia";
        case "TIENDA":
          return defaultValue ? "11200502" : "Bancolombia 1450 Tienda";
        case "SALVIO":
          return defaultValue ? "11200504" : "Bancolombia Salvio 1664";
        case "ACOPIO":
          return defaultValue ? "11200501" : "Bancolombia Ahorros 17800019397";
      }
    } else if (descripcion === "EFECTIVO CAJA") {
      switch (archiveName) {
        case "CHIA":
          return defaultValue ? "11050504" : "Caja General Chia";
        case "TIENDA":
          return defaultValue ? "11050503" : "Caja General Tienda";
        case "SALVIO":
          return defaultValue ? "11050505" : "Caja General Salvio";
        case "P127":
          return defaultValue ? "11050506" : "Caja General P127";
        case "ACOPIO":
          return defaultValue ? "11050502" : "Caja General Planta";
      }
    } else if (descripcion === "PROPINA") {
      switch (archiveName) {
        case "CHIA":
          return defaultValue ? "28150502" : "Propinas";
        case "TIENDA":
          return defaultValue ? "28150502" : "Propinas";
        case "ACOPIO":
          return defaultValue ? "28150503" : "Domicilios";
      }
    } else if (descripcion === "DOMICILIOS: DOMICILIO") {
      return defaultValue ? "28150503" : "Domicilios";
    } else if (descripcion === "CxC RAPPI") {
      return defaultValue ? "13359001" : "Rappi";
    } else if (descripcion === "CxC CLIENTES") {
      return defaultValue ? "13050501" : "Cartera";
    } else if (descripcion === "DESCUENTOS Y CORTESIAS") {
      return defaultValue ? "41400101" : "Descuentos";
    } else {
      return defaultValue ? defaultValue : descripcion;
    }
  }
}
