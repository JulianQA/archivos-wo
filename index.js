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
      value: "FechaElaboracion",
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
    this.tipoDoc = {
      type: String,
      value: this.validationTipoDoc(rows[0], rows[1], rows[20]),
    };
    this.numDoc = { type: Number, value: rows[1] };
    this.fechaElaboracion = { type: String, value: rows[2] };
    this.siglaMoneda = { type: String, value: null };
    this.tasaDeCambio = { type: String, value: null };
    this.vCuentaContable = {
      type: String,
      value: this.validation(rows[19], archiveName),
    };
    this.vNit = { type: String, value: rows[6].toString() };
    this.sucursal = { type: String, value: null };
    this.codProducto = { type: String, value: rows[8] };
    this.bodega = { type: String, value: null };
    this.accion = { type: String, value: null };
    this.cantidadProducto = { type: String, value: rows[11] };
    this.prefijo = { type: String, value: null };
    this.consecutivo = { type: String, value: null };
    this.numCuota = { type: String, value: null };
    this.fechaVencimiento = { type: String, value: null };
    this.codImpuesto = {
      type: String,
      value: this.validationImpuesto(rows[16], rows[19]),
    };
    this.codGrupoActivoFijo = { type: String, value: null };
    this.codActivoFijo = { type: String, value: null };
    this.descripcion = {
      type: String,
      value: this.validation(rows[19], archiveName),
    };
    this.codSubCentroCostos = {
      type: String,
      value: rows[20],
    };
    this.debito = { type: Number, value: rows[21] };
    this.credito = { type: Number, value: rows[22] };
    this.observaciones = { type: String, value: null };
    this.baseGravable = { type: String, value: null };
    this.baseExcenta = { type: String, value: null };
    this.mesCierre = { type: String, value: null };
  }

  validation(descripcion, archiveName) {
    if (descripcion === "TCREDITO" || descripcion === "TDEBITO") {
      if (archiveName === "CHIA") {
        return "13359004";
      } else {
        return "13359003";
      }
    } else if (descripcion === "DAVIVIENDA 3959") {
      return "11200505";
    } else if (descripcion === "WOMPI") {
      return "13359002";
    } else if (descripcion === "Transferencia bancar" || descripcion === "QR") {
      switch (archiveName) {
        case "CHIA":
          return "11200503";
        case "TIENDA":
          return "11200502";
        case "SALVIO":
          return "11200504";
        case "ACOPIO":
          return "11200501";
      }
    } else if (descripcion === "EFECTIVO CAJA") {
      switch (archiveName) {
        case "CHIA":
          return "11050504";
        case "TIENDA":
          return "11050503";
        case "SALVIO":
          return "11050505";
        case "P127":
          return "11050506";
        case "ACOPIO":
          return "11050502";
      }
    } else if (descripcion === "PROPINA") {
      switch (archiveName) {
        case "CHIA":
          return "28150502";
        case "TIENDA":
          return "28150502";
        case "ACOPIO":
          return "28150503";
      }
    } else if (descripcion === "DOMICILIOS: DOMICILIO") {
      return "28150503";
    } else if (descripcion === "CxC RAPPI") {
      return "13359001";
    } else if (
      descripcion === "DEVOLUCION CXC CLIENTES" ||
      descripcion === "Cartera"
    ) {
      return "13050501";
    } else if (descripcion === "DESCUENTOS Y CORTESIAS") {
      return "41400101";
    } else if (descripcion === "ICUI") {
      return "24950301";
    } else {
      return descripcion;
    }
  }

  validationTipoDoc(acc, numDoc, centerCost) {
    if (numDoc.length === 3) {
      return "20";
    } else {
      if (centerCost === "62490-2") {
        return "9";
      } else if (centerCost === "62401-1") {
        return "10";
      }
    }
    return acc;
  }

  validationImpuesto(impuesto, descripcion) {
    if (descripcion === "INC 8%" || descripcion === "DEVOLUCION INC 8%") {
      return "16";
    } else if (descripcion === "ICUI") {
      return "33";
    }

    return impuesto;
  }
}
