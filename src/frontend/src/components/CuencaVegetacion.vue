<template>

<div class="row">
  <div class="column-1" style="background-color: #e3f2fd">
    <MDBListGroup  id="inner" >
      <MDBListGroupItem tag="button" @click="navega('/cuenca-mexico')" action>La cuenca</MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-hidrologia')" action >Hidrología </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-geologia')" action >Geología </MDBListGroupItem>
      <MDBListGroupItem tag="button" action active>Vegetación y uso de suelo</MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-conservacion')" action>Áreas de conservación </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-jurisdiccion')" action >Gestión y marco jurídico</MDBListGroupItem>
    </MDBListGroup>

  </div>
  <div class="column-2" style="background-color:#F8F8F8;">
    <h2>Vegetación y uso de suelo</h2> <br>
    <p align="justify" >La vegetación de la Cuenca de México está representada principalmente por las siguientes comunidades: </p>
    
    <MDBAccordion v-model="activeItem" flush stayOpen>
    
      <MDBAccordionItem
      headerTitle="Bosques"
      headerClasses="acc-head"
      collapseId="flush-collapseOne"
      style="background-color:#F8F8F8;"
    >
      <strong>Bosque de pino (Pinus spp.)</strong>
      <p align="justify"> Típico de los ecosistemas de montaña y crece entre los 2,350 y 4,000 m de altitud, con lluvias anuales de entre 700 y 1,200 mm.  </p>
      <br>
      <strong>Bosque de encino (Quercus spp.)</strong>
      <p align="justify"> Se encuentran entre los 300 y 3,000 m, con lluvias anuales. Con frecuencia los bosques de pino y encino, crecen juntos formando comunidades mixtas. Los encinos tienen entre cinco y 12 m de altura. </p>
      <br>
      <strong>Bosque de enebros (Juniperus spp.)</strong>
      <p align="justify"> Son arbustos o árboles de menos de 6 m de altura, forman bosques bajos y ralos, con abundante vegetación herbácea. Estos bosques son comunes en las partes norte, este y noreste de la cuenca, entre los 2,400 y los 2,800 m de altitud. Se desarrollan sobre laderas o planicies semiáridas, con lluvias anuales de entre 600 y 800 mm. </p>
      </MDBAccordionItem>      
    
    <MDBAccordionItem
      headerTitle="Matorrales"
      collapseId="flush-collapseTwo"
      style="background-color:#F8F8F8;"
    >
      <strong>Matorrales de encinos chaparros (Quercus microphylla)</strong>
      <p align="justify"> Esta comunidad forma una cubierta densa de arbustos bajos (40 a 80 cm de altura). Este tipo de bosque se encuentra principalmente al noreste de la cuenca, en áreas semiáridas con 700 y 900 mm de lluvia media anual. </p>
      <br>
      <strong>Matorrales xerófilos.</strong>
      <p align="justify"> Este tipo de vegetación está compuesta por varias comunidades arbustivas y son frecuentes en la parte norte de la cuenca, donde las precipitaciones son más pobres. En general ocupan partes bajas de la cuenca, entre 2,250 y 2,700 m de altitud, en áreas de precipitación media anual inferior a los 700 mm. En la Sierra de Guadalupe, quedan restos del matorral Eisenhardtia polystachya, que está desapareciendo por la presión antropogénica. </p>
      <br>
      <strong> Vegetación halófila.</strong> 
      <p align="justify"> Esta vegetación domina algunas partes más bajas de la cuenca, sobre los lechos de los antiguos lagos. Es particularmente común el lecho seco del antiguo lago de Texcoco. Este tipo de vegetación ocupaba los márgenes del algo, pero con la construcción del drenaje de la cuenca y el secado de los lagos, se ha extendido a las partes más bajas. </p>
    </MDBAccordionItem>
    <MDBAccordionItem
      headerTitle="Pastizales"
      collapseId="flush-collapseThree"
      style="background-color:#F8F8F8;"
    >
      <strong> Zacate </strong>
      <p align="justify">
      La principal formación de pastizales de la cuenca de México son los de Hilaria cenchroides, comunes al pie de la Sierra Nevada. Esta comunidad se ubica entre los 2,300 y 2,700 m de altitud, con precipitaciones anuales de 600 y 700 mm. </p>
      
      <br>
      <strong> Gramíneas anuales </strong>
      <p align="justify"> En las planicies del centro y norte de la cuenca, a 2,300 y 2,400 m de altitud y sobre áreas fuertemente perturbadas, se desarrolla una comunidad de pastizal en la que dominan Aristida adscensionis y Bouteloua simplex.</p>

      <br>
      <strong> Zacatones </strong>
      <p align="justify">  A más de los 4,000 m, por encima del bosque de Pinus hartwegii, se encuentran los pastizales de Muhlenbergia y Festuca. Estos pastos o zacatones, forman matas erectas de 60 a 120 cm de altura.  </p>
    </MDBAccordionItem>
  </MDBAccordion>
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

.accordion-header{
  background-color: #E3F2FD
}


</style>

<script>
  import { MDBListGroup, MDBListGroupItem, MDBAccordion, MDBAccordionItem } from "mdb-vue-ui-kit";
  import router from '../router'
  export default {
    components: {
      MDBListGroup,
      MDBListGroupItem,
      MDBAccordion,
      MDBAccordionItem,
    }, 
    methods: {
    navega: function(url) {
      router.push(url).catch(()=>{});
      },
    }
  };
</script>