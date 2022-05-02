<template>

<div class="row">
  <div class="column" style="background-color: #e3f2fd">
    <MDBListGroup  id="inner" >
      <MDBListGroupItem tag="button" @click="navega('/cuenca-mexico')" action >La cuenca</MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-hidrologia')" action >Hidrología </MDBListGroupItem>
      <MDBListGroupItem tag="button" action active>Geología </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-vegetacion')" action>Clima y vegetación</MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-conservacion')" action>Áreas de conservación </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-jurisdiccion')" action >Gestión y marco jurídico</MDBListGroupItem>
    </MDBListGroup>

  </div>
  <div class="column-2" style="background-color:#F8F8F8;">
    <h2>Geología de la cuenca</h2>  <br>
    <p align="justify" >La Cuenca de México forma parte del Eje Neovolcánico o Faja Volcánica Transmexicana, la cual se extiende en dirección Este-Oeste, atravesando la República Mexicana. Esta zona se considera una provincia fisiográfica independiente al resto de la República y se encuentra limitada por elevaciones topográficas de origen volcánico como la Sierra Guadalupe al norte, la sierra de las Cruces al poniente, la Sierra del Chichinautzin al sur y la Sierra Nevada al oriente.</p>
    <p align="justify">En relación al tipo de sedimento, se caracteriza por una composición de materiales volcánicos, intercalados con aluviones y cubiertos, en la parte central del valle, por arcillas lacustres. La mayor parte de las rocas y materiales que forman el subsuelo son permeables.</p>
    <img src="../assets/geologia.png" height="400" width="610" class="img-fluid  hover-shadow" alt="..." />
    <br> <br>
    <p align="justify">Los materiales que constituyen el subsuelo corresponden a una intercalación de productos volcánicos como lavas, tobas y ceniza que incluyen materiales granulares transportados por ríos y arroyos provenientes de las partes topográficamente altas que circularon hacia los valles.</p>
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

/* Create two columns that sits next to each other */
.column {
  float: left;
  width: 25%;
  padding: 10px;
  padding-top: 25px;
}
.column-2 {
  float: left;
  width: 75%;
  padding: 10px;
}

/* Clear floats after the columns */
.row:after {
  content: "";
  display: table;
  clear: both;
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