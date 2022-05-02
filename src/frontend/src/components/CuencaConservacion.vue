<template>

<div class="row">
  <div class="column" style="background-color:#e3f2fd;">
    <MDBListGroup  id="inner" >
      <MDBListGroupItem tag="button" @click="navega('/cuenca-mexico')" action >La cuenca</MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-hidrologia')" action >Hidrología </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-geologia')" action >Geología </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-vegetacion')" action>Clima y vegetación</MDBListGroupItem>
      <MDBListGroupItem tag="button"  action active>Áreas de conservación </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-jurisdiccion')" action >Gestión y marco jurídico</MDBListGroupItem>
    </MDBListGroup> 
  </div>
  <div class="column-2" style="background-color:#F8F8F8;">
    <h2>Áreas de conservación </h2>  <br>
    <p align="justify" >La Ciudad de México está dividido en dos zonas administrativas, el suelo de conservación y la zona urbana. Las subcuencas ubicadas en ésta zona son ríos La Magdalena-Eslava, Santo Desierto y Las Regaderas. </p>
    <p align="justify" > La cuenca conforma un paisaje diverso y rico en el que, consecuentemente, se presenta un complejo de climas, suelos y vegetación local que debe ser conservado.</p>
    <div class="bg-image hover-zoom"> 
      <img src="../assets/conservacion.png" height="510" width="630" class="img-fluid  hover-shadow" alt="..." />
    </div> 
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