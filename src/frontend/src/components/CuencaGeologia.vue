<template>

<div class="row">
  <div class="column-1" style="background-color: #e3f2fd">
    <MDBListGroup  id="inner" >
      <MDBListGroupItem tag="button" @click="navega('/cuenca-mexico')" action >La cuenca</MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-hidrologia')" action >Hidrología </MDBListGroupItem>
      <MDBListGroupItem tag="button" action active>Geología </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-vegetacion')" action>Vegetación y uso de suelo</MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-conservacion')" action>Áreas de conservación </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-jurisdiccion')" action >Gestión y marco jurídico</MDBListGroupItem>
    </MDBListGroup>

  </div>
  <div class="column-2" style="background-color:#F8F8F8;">
    <h2>Geología de la cuenca</h2>  <br>
    <p align="justify" >La Cuenca de México forma parte del Eje Neovolcánico o Faja Volcánica Transmexicana, la cual se extiende en dirección Este-Oeste, atravesando la República Mexicana. Esta zona se considera una provincia fisiográfica independiente al resto de la República y se encuentra limitada por elevaciones topográficas de origen volcánico como la Sierra Guadalupe al norte, la sierra de las Cruces al poniente, la Sierra del Chichinautzin al sur y la Sierra Nevada al oriente.</p>
    <p align="justify">En relación al tipo de sedimento, se caracteriza por una composición de materiales volcánicos, intercalados con aluviones y cubiertos, en la parte central del valle, por arcillas lacustres. La mayor parte de las rocas y materiales que forman el subsuelo son permeables.</p>
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

/* Create two equal columns that sits next to each other */
.column-1 {
  flex: 25%;
  padding-top: 30px;
  /*padding-left: 50px;*/
  min-height: 100px;
  overflow: hidden;
}
.column-2 {
  flex: 75%;
  padding: 40px;
  min-height: 100px;
  overflow: hidden;
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