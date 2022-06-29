<template>

<div class="row">
  <div class="column" style="background-color: #e3f2fd">
    <MDBListGroup  id="inner" >
      <MDBListGroupItem tag="button" action active>Protocolos</MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/protocolo-fisicoquimica')" action > Calidad fisicoquímica </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-geologia')" action > HYQI </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-vegetacion')" action>MEQI</MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-vegetacion')" action>DEQI</MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-conservacion')" action>BEMEQI </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-jurisdiccion')" action >Material de trabajo</MDBListGroupItem>
    </MDBListGroup>

  </div>
  <div class="column-2" style="background-color:#F8F8F8;">
    <h2>Protocolos de calidad ecológica</h2> <br>
    <img src="../assets/bajo_agua.jpg" width="215" class="img-fluid  hover-shadow" alt="..." 
      style="float:right;
      margin-left: 20px" />
    <p align="justify" >El impacto de la actividad humana en los sistemas acuáticos es evaluado principalmente a través de la calidad del agua, la cual determina su naturaleza química, física y biológica, así como su relación con el uso y los efectos en la salud humana. </p>
    <br>
    <p  align="justify" > Otro componente que ha tomado mayor relevancia en los últimos años, es la evaluación del ecosistema a través de las comunidades biológicas llamado biomonitoreo, el cual intenta evaluar y dar seguimiento a los cambios espaciales y temporales en el ambiente. Sin embargo, la integración de ambas aproximaciones ha sido controversial ya que se ponderan con distinto valor los parámetros más adecuados para su evaluación y monitoreo. </p>
    <br>
    <img src="../assets/estudio.jpg"  width="220" class="img-fluid  hover-shadow" alt="..." 
      style="float:left;
      margin-right: 20px" />
    <p align="justify"> Una propuesta que vincula ambas aproximaciones y que se emplea en la presente investigación, es la integridad del ecosistema, la cual implica seleccionar los sitios de referencia, los parámetros fisicoquímicos y organismos indicadores que reflejen el estado y el grado de deterioro del ecosistema de ribera en el tiempo y el espacio. Esta propuesta introduce una revisión teórico-metodológica que parte de la premisa de que el beneficio de la provisión del agua, responde a una serie de procesos que reflejan el estado del funcionamiento del ecosistema de manera integral. Así mismo, se plantean las herramientas útiles para la construcción de índices bióticos a partir de enfoques multimétricos.</p>
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
  padding: 40px;
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