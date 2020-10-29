pipeline {
  agent {
    node {
      label 'node1'
    }

  }
  stages {
    stage('Stage1') {
      steps {
        sh 'dotnet build'
      }
    }

  }
}