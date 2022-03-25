<template>

<div class="row">
  <div class="column-1" style="background-color: #e3f2fd">
    <MDBListGroup  id="inner" >
      <MDBListGroupItem tag="button" @click="navega('/cuenca-mexico')" action >La cuenca</MDBListGroupItem>
      <MDBListGroupItem tag="button" action active>Hidrología </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-geologia')" action >Geología </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-vegetacion')" action>Vegetación y uso de suelo</MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-conservacion')" action>Áreas de conservación </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-jurisdiccion')" action >Gestión y marco jurídico</MDBListGroupItem>
    </MDBListGroup>

  </div>
  <div class="column-2" style="background-color:#F8F8F8;">
    <h2>Hidrología de la cuenca</h2>  <br>
    <p align="justify" >La cuenca está dividida políticamente en cinco estados de la República Mexicana(Estado de México, Ciudad de México, Tlaxcala, Hidalgo y Puebla) y se caracteriza por ser una zona montañosa que forma numerosos caudales que drenan hacia el valle.</p>
    <p align="justify">Se encuentra en la región morfotectónica de la Faja Volcánica Transmexicana y tiene una superficie total de 9,600 km2, de los cuales 5,518 km2 están en áreas montañosas por encima de los 2,400 metros sobre el nivel del mar. </p>
    <p align="justify">El clima de la región es templado subhúmedo (temperatura media anual de 13.4 ºC, la precipitación media anual de entre 1,200 a 1,500 mm). De forma natural, en la cuenca se puede reconocer en los ciclos hidrológicos una época de lluvias (junio a noviembre), otra época de secas frías (diciembre a febrero) y secas cálidas (marzo a mayo); condición que ha sufrido poca modificación a lo largo del tiempo. </p>
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