const archiveInput = document.querySelector(".form__archive");

archiveInput.addEventListener("change", async () => {
  const response = await readXlsxFile(archiveInput.files[0]);
  const nameArchive = archiveInput.files[0].name;
  const list = nameArchive.split(" ");
  const woName = `ARCHIVO WO ${list[1]} ${list[2]} ${list[3]} ${
    list[list.length - 1].split(".")[0]
  }`;
  const excel = new Excel(response).getRow();
  excel.unshift([
    {
      value: "Encab: Empresa",
      fontWeight: "bold",
    },
    {
      value: "Encab: Tipo Documento",
      fontWeight: "bold",
    },
    {
      value: "Encab: Prefijo",
      fontWeight: "bold",
    },
    {
      value: "Encab: Documento Número",
      fontWeight: "bold",
    },
    {
      value: "Encab: Fecha",
      fontWeight: "bold",
    },
    {
      value: "Encab: Tercero Interno",
      fontWeight: "bold",
    },
    {
      value: "Encab: Tercero Externo",
      fontWeight: "bold",
    },
    {
      value: "Encab: Nota",
      fontWeight: "bold",
    },
    {
      value: "Encab: FormaPago",
      fontWeight: "bold",
    },
    {
      value: "Encab: Fecha Entrega",
      fontWeight: "bold",
    },
    {
      value: "Encab: Prefijo Documento Externo",
      fontWeight: "bold",
    },
    {
      value: "Encab: Número_Documento_Externo",
      fontWeight: "bold",
    },
    {
      value: "Encab: Verificado",
      fontWeight: "bold",
    },
    {
      value: "Encab: Anulado",
      fontWeight: "bold",
    },
    {
      value: "Encab: Personalizado 1",
      fontWeight: "bold",
    },
    {
      value: "Encab: Personalizado 2",
      fontWeight: "bold",
    },
    {
      value: "Encab: Personalizado 3",
      fontWeight: "bold",
    },
    {
      value: "Encab: Personalizado 4",
      fontWeight: "bold",
    },
    {
      value: "Encab: Personalizado 5",
      fontWeight: "bold",
    },
    {
      value: "Encab: Personalizado 6",
      fontWeight: "bold",
    },
    {
      value: "Encab: Personalizado 7",
      fontWeight: "bold",
    },
    {
      value: "Encab: Personalizado 8",
      fontWeight: "bold",
    },
    {
      value: "Encab: Personalizado 9",
      fontWeight: "bold",
    },
    {
      value: "Encab: Personalizado 10",
      fontWeight: "bold",
    },
    {
      value: "Encab: Personalizado 11",
      fontWeight: "bold",
    },
    {
      value: "Encab: Personalizado 12",
      fontWeight: "bold",
    },
    {
      value: "Encab: Personalizado 13",
      fontWeight: "bold",
    },
    {
      value: "Encab: Personalizado 14",
      fontWeight: "bold",
    },
    {
      value: "Encab: Personalizado 15",
      fontWeight: "bold",
    },
    {
      value: "Encab: Sucursal",
      fontWeight: "bold",
    },
    {
      value: "Encab: Clasificación",
      fontWeight: "bold",
    },
    {
      value: "Detalle: Producto",
      fontWeight: "bold",
    },
    {
      value: "Detalle: Bodega",
      fontWeight: "bold",
    },
    {
      value: "Detalle: UnidadDeMedida",
      fontWeight: "bold",
    },
    {
      value: "Detalle: Cantidad",
      fontWeight: "bold",
    },
    {
      value: "Detalle: IVA",
      fontWeight: "bold",
    },
    {
      value: "Detalle: Valor Unitario",
      fontWeight: "bold",
    },
    {
      value: "Detalle: Descuento",
      fontWeight: "bold",
    },
    {
      value: "Detalle: Vencimiento",
      fontWeight: "bold",
    },
    {
      value: "Detalle: Nota",
      fontWeight: "bold",
    },
    {
      value: "Detalle: Centro costos",
      fontWeight: "bold",
    },
    {
      value: "Detalle: Personalizado1",
      fontWeight: "bold",
    },
    {
      value: "Detalle: Personalizado2",
      fontWeight: "bold",
    },
    {
      value: "Detalle: Personalizado3",
      fontWeight: "bold",
    },
    {
      value: "Detalle: Personalizado4",
      fontWeight: "bold",
    },
    {
      value: "Detalle: Personalizado5",
      fontWeight: "bold",
    },
    {
      value: "Detalle: Personalizado6",
      fontWeight: "bold",
    },
    {
      value: "Detalle: Personalizado7",
      fontWeight: "bold",
    },
    {
      value: "Detalle: Personalizado8",
      fontWeight: "bold",
    },
    {
      value: "Detalle: Personalizado9",
      fontWeight: "bold",
    },
    {
      value: "Detalle: Personalizado10",
      fontWeight: "bold",
    },
    {
      value: "Detalle: Personalizado11",
      fontWeight: "bold",
    },
    {
      value: "Detalle: Personalizado12",
      fontWeight: "bold",
    },
    {
      value: "Detalle: Personalizado13",
      fontWeight: "bold",
    },
    {
      value: "Detalle: Personalizado14",
      fontWeight: "bold",
    },
    {
      value: "Detalle: Personalizado15",
      fontWeight: "bold",
    },
    {
      value: "Detalle: Código Centro Costos",
      fontWeight: "bold",
    },
  ]);
  return await writeXlsxFile(excel, {
    fileName: woName,
  });
});

class Excel {
  constructor(archive) {
    this.archive = archive;
  }

  getRow() {
    let rows = [];
    let anterior;
    for (let i = 1; i < this.archive.length; i++) {
      let actual = Object.values(new Rows(this.archive[i]));
      if (actual[8].value === "Vacío") {
        actual[8].value = anterior[8].value;
      }
      rows.push(actual);
      anterior = actual;
    }
    return rows;
  }
}

class Rows {
  constructor(rows) {
    this.empresa = { type: String, value: rows[0] };
    this.tipoDocumento = { type: String, value: "FV" };
    this.prefijo = { type: String, value: rows[2] };
    this.documentoNumero = { type: Number, value: rows[3] };
    this.fecha = { type: String, value: rows[4] };
    this.terceroInterno = {
      type: Number,
      value: this.validationVendedor(rows[2]),
    };
    this.terceroExterno = { type: Number, value: 222222222 };
    this.nota = { type: String, value: "Factura de Venta" };
    this.formaPago = {
      type: String,
      value: this.validationFormaPago(rows[9], rows[2]),
    };
    this.fechaEntrega = { type: String, value: rows[4] };
    this.prefijoDocumentoExterno = { type: String, value: null };
    this.numeroDocumentoExterno = { type: String, value: null };
    this.verificado = { type: Number, value: -1 };
    this.anulado = { type: Number, value: 0 };
    this.presonalizado1 = { type: String, value: null };
    this.presonalizado2 = { type: String, value: null };
    this.presonalizado3 = { type: String, value: null };
    this.presonalizado4 = { type: String, value: null };
    this.presonalizado5 = { type: String, value: null };
    this.presonalizado6 = { type: String, value: null };
    this.presonalizado7 = { type: String, value: null };
    this.presonalizado8 = { type: String, value: null };
    this.presonalizado9 = { type: String, value: null };
    this.presonalizado10 = { type: String, value: null };
    this.presonalizado11 = { type: String, value: null };
    this.presonalizado12 = { type: String, value: null };
    this.presonalizado13 = { type: String, value: null };
    this.presonalizado14 = { type: String, value: null };
    this.presonalizado15 = { type: String, value: null };
    this.surcursal = { type: String, value: null };
    this.clasificacion = { type: String, value: null };
    this.dettaleProducto = {
      type: String,
      value: rows[10] === "Rappi Valor Variable" ? "PT298" : rows[10],
    };
    this.bodega = { type: String, value: "LUISA POSTRES" };
    this.unidadDeMedida = { type: String, value: "Und." };
    this.cantidad = { type: Number, value: rows[11] };
    this.iva = { type: Number, value: 0 };
    this.valorUnitario = { type: Number, value: rows[13] };
    this.descuento = { type: Number, value: rows[18] / (rows[11] * rows[13]) };
    this.fechaVencimiento = { type: String, value: rows[4] };
    this.detalleNota = { type: String, value: null };
    this.centroCostos = {
      type: String,
      value: this.validationCentroCostos(rows[2]),
    };
    this.detallePer1 = { type: String, value: null };
    this.detallePer2 = { type: String, value: null };
    this.detallePer3 = { type: String, value: null };
    this.detallePer4 = { type: String, value: null };
    this.detallePer5 = { type: String, value: null };
    this.detallePer6 = { type: String, value: null };
    this.detallePer7 = { type: String, value: null };
    this.detallePer8 = { type: String, value: null };
    this.detallePer9 = { type: String, value: null };
    this.detallePer10 = { type: String, value: null };
    this.detallePer11 = { type: String, value: null };
    this.detallePer12 = { type: String, value: null };
    this.detallePer13 = { type: String, value: null };
    this.detallePer14 = { type: String, value: null };
    this.detallePer15 = { type: String, value: null };
    this.codCentroCostos = { type: String, value: null };
  }
  getFormaPago() {
    return this.formaPago;
  }
  setFormaPago(formaPago) {
    this.formaPago = formaPago;
  }
  validationFormaPago(element, prefijo) {
    if (element === "TCREDITO" || element === "TDEBITO") {
      return "Link de pago bold";
    } else if (element === "EFECTIVO CAJA") {
      return "Contado";
    } else if (element === "CUENTAS x COBRAR") {
      return "RAPPI";
    } else if (element === "OTROS") {
      if (prefijo === "CH") {
        return "Bancolombia 1663 Chia";
      } else if (prefijo === "FIM") {
        return "Bancolombia 1664 Salvio";
      } else if (prefijo === "TB") {
        return "Bancolombia 1450";
      }
    } else {
      return "Vacío";
    }
  }
  validationCentroCostos(prefijo) {
    if (prefijo === "CH") {
      return "Chia";
    } else if (prefijo === "FIM") {
      return "Salvio";
    } else if (prefijo === "TB") {
      return "Tienda";
    } else if (prefijo === "P127") {
      return "127";
    } else {
      return "Planta";
    }
  }
  validationVendedor(prefijo) {
    if (prefijo === "CH") {
      return 52267032;
    } else if (prefijo === "FIM") {
      return 1069505011;
    } else if (prefijo === "TB") {
      return 1094902055;
    } else if (prefijo === "P127") {
      return 39678885;
    } else {
      return 1000236135;
    }
  }
}
