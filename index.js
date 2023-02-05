const archiveInput = document.querySelector(".form__archive");

archiveInput.addEventListener("change", async () => {
  const response = await readXlsxFile(archiveInput.files[0]);
  const nameArchiveRead = archiveInput.files[0].name;
  const name = nameArchiveRead.split(" ");
  const woName = `ARCHIVO WO ${name[0]} ${name[2]} ${name[3]} ${
    name[4].split(".")[0]
  }`;
  console.log(response, woName);
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
      value: this.validationCuentaContable(rows[19], archiveName, rows[5]),
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
      value: this.validationDescripcion(rows[19], archiveName),
    };
    this.codSubCentrocostos = { type: String, value: null };
    this.debito = { type: Number, value: rows[21] };
    this.credito = { type: Number, value: rows[22] };
    this.observaciones = { type: String, value: null };
    this.baseGravable = { type: String, value: null };
    this.baseExcenta = { type: String, value: null };
    this.mesCierre = { type: String, value: null };
  }
  validationDescripcion(descripcion, archiveName) {
    if (descripcion === "TCREDITO" || descripcion === "TDEBITO") {
      if (archiveName === "CHIA") {
        return "Bold";
      } else {
        return "Credibanco";
      }
    } else if (descripcion === "DAVIVIENDA 3959") {
      return "Davivienda 3959";
    } else if (descripcion === "WOMPI") {
      return "Wompi";
    } else if (descripcion === "Transferencia bancar") {
      switch (archiveName) {
        case "CHIA":
          return "Bancolombia 1663 Chia";
        case "TIENDA":
          return "Bancolombia 1450 Tienda";
        case "SALVIO":
          return "Bancolombia Salvio 1664";
        case "ACOPIO":
          return "Bancolombia Ahorros 17800019397";
      }
    } else if (descripcion === "EFECTIVO CAJA") {
      switch (archiveName) {
        case "CHIA":
          return "Caja General Chia";
        case "TIENDA":
          return "Caja General Tienda";
        case "SALVIO":
          return "Caja General Salvio";
        case "P127":
          return "Caja General P127";
        case "ACOPIO":
          return "Caja General Planta";
      }
    } else if (descripcion === "PROPINA") {
      switch (archiveName) {
        case "CHIA":
          return "Propinas";
        case "TIENDA":
          return "Propinas";
        case "ACOPIO":
          return "Domicilios";
      }
    } else if (descripcion === "DOMICILIOS: DOMICILIO") {
      return "Domicilios";
    } else if (descripcion === "CxC RAPPI") {
      return "Rappi";
    } else {
      return descripcion;
    }
  }
  validationCuentaContable(descripcion, archiveName, defaultValue) {
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
    } else if (descripcion === "Transferencia bancar") {
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
      return "Domicilios";
    } else if (descripcion === "CxC RAPPI") {
      return "13359001";
    } else {
      return defaultValue;
    }
  }
}
