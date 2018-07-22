pipeline {
    agent any
    stages {
        stage('Test') {
            steps {
                echo "Test"
            }
        }
    }
    post {
        always {
            junit 'build/reports/**/*.xml'
        }
    }
}
