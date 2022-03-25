<template>

<div class="row">
  <div class="column-1" style="background-color: #e3f2fd">
    <MDBListGroup  id="inner" >
      <MDBListGroupItem tag="button" action active>La cuenca</MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-hidrologia')" action >Hidrología </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-geologia')" action >Geología </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-vegetacion')" action>Vegetación y uso de suelo</MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-conservacion')" action>Áreas de conservación </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-jurisdiccion')" action >Gestión y marco jurídico</MDBListGroupItem>
    </MDBListGroup>

  </div>
  <div class="column-2" style="background-color:#F8F8F8;">
    <h2>La cuenca de México</h2> <br>
    <p align="justify" >La cuenca de México pertenece a la región hidrológica no. 26 y hasta el año 2009, se reportan 45 ríos vivos que nacen en las zonas motañosas que delimitan a la cuenca.  En la cuenca de México se ha estudiado la estructura de la comunidad de macroalgas, diatomeas, macroinvertebrados y vegetación vascular ribereña a lo largo de 8 años continuos.  </p>
    <p align="justify">Por una lado, se ha logrado establecer las condiciones de referencia que caracterizan a los ríos de montaña que mantienen un elevado estatus de conservación y aquellos sitios que han perdido la buena condición por alguna acción humana. Y por el otro, el análisis de la diversidad acuática y su relación con la calidad fisicoquímica e hidromorfológica ha permitido asignarle un valor indicador a cada taxa y la propuesta de un índice relacionado con clases de calidad, que en conjunto, son una propuesta que permitirá evaluar la calidad ecológica en los ríos de la cuenca de México</p>
    <p align="justify">En México, la incipiente interacción e interés entre la ciencia y las autoridades que administran las aguas nacionales, explica en gran medida, la ausencia de leyes que promuevan su uso para su monitoreo y conservación. De tal manera que, la integración entre la composición y funcionamiento de la diversidad biológica y los agentes naturales y/o antrópicos que la modifican, permiten establecer un valor indicador de conservación del ecosistema. Este valor indicador, se integra en un análisis que integra a la comunidad acuática y permite monitorear con un algoritmo, relativamente sencillo de utilizar, el estado de calidad ecológica. </p>
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