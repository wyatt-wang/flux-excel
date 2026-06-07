package com.github.export;

import lombok.extern.slf4j.Slf4j;
import org.junit.Test;

import java.io.ByteArrayInputStream;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 线程测试
 */
@Slf4j
public class ThreadTest {
	@Test
	public void test1()throws Exception{
		ThreadJob threadJob = new ThreadJob();
		for (int i = 1; i <= 20; i++) {
			Thread thread = new Thread(threadJob,"thread-"+i);
			thread.start();
		}
		Thread.sleep(3 * 60 * 1000);
	}

	class ThreadJob implements Runnable {
		byte[] bytes = new byte[2048];
		ThreadLocal<ByteArrayInputStream> threadLocal = ThreadLocal.withInitial(()->{
			System.out.println(bytes.toString());
			return new ByteArrayInputStream(bytes);
		}) ;
		@Override
		public void run() {
			while (true) {
				ByteArrayInputStream inputStream = threadLocal.get();
				try {
					Thread.sleep(3 * 60 * 1000);
				} catch (InterruptedException e) {
					e.printStackTrace();
				}
			}
			/*Scanner scanner = new Scanner(System.in);
			String str = scanner.nextLine();*/
		}
	}

}
