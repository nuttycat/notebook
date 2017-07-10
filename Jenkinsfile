node{
stage('mvn test'){
        //mvn 测试
        echo "mvn test"
    }

    stage('mvn build'){
        //mvn构建
        echo "mvn clean install -Dmaven.test.skip=true"
    }

    stage('deploy'){
        //执行部署脚本
        echo "deploy ......" 
    }
}
