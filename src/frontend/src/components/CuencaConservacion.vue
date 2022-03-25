<template>

<div class="row">
  <div class="column-1" style="background-color: #e3f2fd">
    <MDBListGroup  id="inner" >
      <MDBListGroupItem tag="button" @click="navega('/cuenca-mexico')" action >La cuenca</MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-hidrologia')" action >Hidrología </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-geologia')" action >Geología </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-vegetacion')" action>Vegetación y uso de suelo</MDBListGroupItem>
      <MDBListGroupItem tag="button"  action active>Áreas de conservación </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-jurisdiccion')" action >Gestión y marco jurídico</MDBListGroupItem>
    </MDBListGroup>

  </div>
  <div class="column-2" style="background-color:#F8F8F8;">
    <h2>Áreas de conservación </h2>  <br>
    <p align="justify" >La Ciudad de México está dividido en dos zonas administrativas, el suelo de conservación y la zona urbana.</p>
    <br>
    <p align="justify" > Las subcuencas ubicadas en ésta zona son ríos La Magdalena-Eslava, Santo Desierto y Las Regaderas. </p>
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
  padding-bottom: 30px;
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