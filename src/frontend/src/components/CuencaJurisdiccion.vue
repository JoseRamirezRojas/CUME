<template>

<div class="row">
  <div class="column" style="background-color: #e3f2fd">
    <MDBListGroup  id="inner" >
      <MDBListGroupItem tag="button" @click="navega('/cuenca-mexico')" action >La cuenca</MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-hidrologia')" action >Hidrología </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-geologia')" action >Geología </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-vegetacion')" action>Clima y vegetación</MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-conservacion')" action >Áreas de conservación </MDBListGroupItem>
      <MDBListGroupItem tag="button" action active>Gestión y marco jurídico</MDBListGroupItem>
    </MDBListGroup>

  </div>
  <div class="column-2" style="background-color:#F8F8F8;">
    <h2>Gestión jurídica</h2>  <br>
    <p align="justify" > En el mundo y a nivel nacional  existen protocolos estandarizados de evaluación, manejo y control de calidad para la conservación de los ríos, algunos de estos pueden visualizarse entre los <a href="#footer"> sitios de interés </a> de esta página. Con las directrices claras y en funcionamiento, se establece un marco legal que los usuarios deben cumplir para evitar sanciones económicas ó la cárcel.  </p>
    <br>
    <p align="justify" >De manera fundamental, la gestión jurídica del agua en México tiene como base lo que dictan tres artículos de la Constitución Política (4˚, 27° y 115°) y la <a href="https://www.diputados.gob.mx/LeyesBiblio/pdf/LAN.pdf" target="_blank" rel="noopener noreferrer" > Ley de Aguas Nacionales </a> (LAN). En la legislación secundaria, la Ley de Aguas Nacionales regula la distribución y control del agua, y designa a la Comisión Nacional del Agua como el órgano responsable de ejercer la autoridad y administración del agua a nombre del Ejecutivo. La LAN se promulgó en 1992 y su última reforma fue publicada en el Diario Oficial de la Federación en marzo del 2016. </p>
    <br>
    <p align="justify" > <a href="../assets/files/NOM_001_002_003_SEMARNAT.pdf " target="_blank">Las Normas oficiales para la calidad del agua en México </a>  
      están regidas por la <a href="assets/files/nom-127-ssai.pdf " target="_blank">NOM-127-SSA1-1994 </a> (DOF 1994) que excluye los componentes biológicos y ambientales en sus variables de control sanitario. 
      No obstante, en el año 2012 se planteó el proyecto de ley <a href="assets/files/NMX-AA-159-SCFI-2012.pdf " target="_blank">Norma Oficial NMX-AA-159-SCFI-2012 </a> que establece los procedimientos para la determinación del caudal ecológico y la calidad del hábitat fluvial con base en indicadores ecológicos y de gestión del agua que sean oportunamente identificados en cuencas hidrológicas.
    </p>
    <br>
    
  </div>
</div>

</template>

<style>
* {
  box-sizing: border-box;
}

.row {
  padding-top: 50px;
  display: flex;
}

/* Create two equal columns that sits next to each other */
.column {
  float: left;
  width: 25%;
  padding: 10px;
  padding-top: 25px;
}
.column-2 {
  flex: 75%;
  padding: 40px;
}

/* Responsive layout - makes the two columns stack on top of each other instead of next to each other */
@media screen and (max-width: 600px) {
  .column {
    width: 100%;
  }
  .column-2 {
    width: 100%;
  }
}

#inner {
  display: table;
  margin: 0 auto;
}


</style>

<script>
  import { MDBListGroup, MDBListGroupItem } from "mdb-vue-ui-kit";
  import router from '../router'
  export default {
    components: {
      MDBListGroup,
      MDBListGroupItem
    },
    methods: {
    navega: function(url) {
      router.push(url).catch(()=>{});
      },
    }
  };
</script>
-->